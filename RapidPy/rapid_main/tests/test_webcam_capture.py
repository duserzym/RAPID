"""Exercise camera preview and snapshot behavior with synthetic Qt frames."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from PySide6 import QtCore, QtGui, QtMultimedia, QtTest, QtWidgets

from webcam_viewer.app import WebcamWidget


class WebcamCaptureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def setUp(self):
        self.widget = WebcamWidget()
        self.addCleanup(self.widget.deleteLater)
        self.addCleanup(self.widget.disconnect_camera)

    def test_preview_and_snapshot_use_the_same_oriented_frame(self):
        image = QtGui.QImage(3, 2, QtGui.QImage.Format.Format_RGB32)
        image.fill(QtGui.QColor("black"))
        image.setPixelColor(0, 0, QtGui.QColor("red"))
        frame = QtMultimedia.QVideoFrame(image)
        frame.setRotation(QtMultimedia.QtVideo.Rotation.Clockwise90)
        frame.setMirrored(True)
        sink = self.widget._begin_connection()
        self.assertFalse(self.widget._snapshot_btn.isEnabled())
        sink.setVideoFrame(frame)
        self.qt.processEvents()
        self.assertTrue(self.widget._connected)
        self.assertTrue(self.widget._snapshot_btn.isEnabled())
        with TemporaryDirectory() as temporary:
            self.widget._snapshot_dir = Path(temporary)
            self.widget._snapshot_btn.click()
            saved, = self.widget._snapshot_dir.glob("*.png")
            snapshot = QtGui.QImage(str(saved))
            self.assertEqual(snapshot.size(), QtCore.QSize(2, 3))
            self.assertEqual(snapshot.pixelColor(0, 0), QtGui.QColor("red"))
        self.widget.disconnect_camera()
        self.assertFalse(self.widget._snapshot_btn.isEnabled())
        self.assertTrue(self.widget._image.isNull())

    def test_old_frames_cannot_reconnect_a_changed_source(self):
        image = QtGui.QImage(10, 10, QtGui.QImage.Format.Format_RGB32)
        image.fill(QtGui.QColor("blue"))
        first = self.widget._begin_connection()
        first.setVideoFrame(QtMultimedia.QVideoFrame(image))
        current = self.widget._begin_connection()
        self.qt.processEvents()
        self.assertIs(self.widget._sink, current)
        self.assertFalse(self.widget._connected)
        self.assertFalse(self.widget._snapshot_btn.isEnabled())
        current.setVideoFrame(QtMultimedia.QVideoFrame(image))
        self.qt.processEvents()
        self.assertTrue(self.widget._connected)

    def test_first_frame_timeout_restores_connect_button(self):
        self.widget._frame_deadline.setInterval(10)
        self.widget._begin_connection()
        for _ in range(20):
            if self.widget._sink is None:
                break
            QtTest.QTest.qWait(10)
        self.assertIsNone(self.widget._sink)
        self.assertEqual(self.widget._connect_btn.text(), "Connect")
        self.assertFalse(self.widget._snapshot_btn.isEnabled())

    def test_network_input_uses_qt_playback_only_for_camera_urls(self):
        self.widget._cam_combo.setCurrentIndex(self.widget._cam_combo.findData("ip"))
        with patch.object(self.widget, "_open_stream") as player:
            self.widget._url_edit.setText("file:///C:/local.avi")
            self.widget._connect_btn.click()
            player.assert_not_called()
            self.widget._url_edit.setText("rtsp://camera.example/stream")
            self.widget._connect_btn.click()
            url = player.call_args.args[0]
            self.assertEqual((url.scheme(), url.host(), url.path()), ("rtsp", "camera.example", "/stream"))

    def test_usb_selection_preserves_device_id_across_reorder_and_unplug(self):
        class Device:
            def __init__(self, identity):
                self.identity = identity

            def id(self):
                return self.identity

            def description(self):
                return "USB Camera"

        first, selected = Device(b"camera-1"), Device(b"camera-2")
        camera, session = Mock(), Mock()
        with (patch.object(QtMultimedia, "QCameraDevice", Device),
              patch.object(QtMultimedia, "QCamera", return_value=camera) as open_camera,
              patch.object(QtMultimedia, "QMediaCaptureSession", return_value=session),
              patch.object(self.widget._devices, "videoInputs", return_value=[first, selected]) as inputs):
            self.widget._populate_cameras()
            self.widget._cam_combo.setCurrentIndex(1)
            inputs.return_value = [selected, first]
            self.widget._populate_cameras()
            self.widget._connect_btn.click()
            open_camera.assert_called_once_with(selected, self.widget)
            session.setVideoSink.assert_called_once_with(self.widget._sink)
            camera.start.assert_called_once()
            inputs.return_value = [first]
            self.widget._populate_cameras()
            camera.stop.assert_called_once()
            self.assertIsNone(self.widget._sink)
            self.assertEqual(self.widget._status_label.text(), "USB camera disconnected")
