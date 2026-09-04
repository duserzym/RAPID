#!/usr/bin/env bash
# Compile aux_test.ino against the bundled Arduino 1.0.5 core and flash an Uno.
# The installed IDE predates CLI upload support, so this drives avr-gcc and
# avrdude directly with the same recipe boards.txt describes.
set -euo pipefail

A="/c/Program Files (x86)/Arduino"
CORE="$A/hardware/arduino/cores/arduino"
VARIANT="$A/hardware/arduino/variants/standard"
BIN="$A/hardware/tools/avr/bin"
CONF="$A/hardware/tools/avr/etc/avrdude.conf"

SRC="/e/Github/RAPID/hardware/arduino-aux-test/aux_test.ino"
OUT="$(dirname "$0")/build"
PORT="${1:-COM13}"

MCU=atmega328p
F_CPU=16000000L
DEFS="-mmcu=$MCU -DF_CPU=$F_CPU -DARDUINO=105"
WARN="-w"
OPT="-Os -ffunction-sections -fdata-sections -MMD"

rm -rf "$OUT"; mkdir -p "$OUT/core"

echo "== core =="
for f in "$CORE"/*.c; do
  "$BIN/avr-gcc" -c $OPT $DEFS $WARN -std=gnu99 -I"$CORE" -I"$VARIANT" "$f" -o "$OUT/core/$(basename "$f").o"
done
for f in "$CORE"/*.cpp; do
  "$BIN/avr-g++" -c $OPT $DEFS $WARN -fno-exceptions -I"$CORE" -I"$VARIANT" "$f" -o "$OUT/core/$(basename "$f").o"
done
"$BIN/avr-ar" rcs "$OUT/core.a" "$OUT"/core/*.o

echo "== sketch =="
# A .ino is C++ with an implicit Arduino.h. Every function here is defined
# before use, so no prototype generation is needed -- a plain include is enough.
{ echo '#include <Arduino.h>'; cat "$SRC"; } > "$OUT/sketch.cpp"
"$BIN/avr-g++" -c $OPT $DEFS $WARN -fno-exceptions -I"$CORE" -I"$VARIANT" "$OUT/sketch.cpp" -o "$OUT/sketch.o"

echo "== link =="
"$BIN/avr-gcc" -Os -Wl,--gc-sections -mmcu=$MCU -o "$OUT/sketch.elf" \
  "$OUT/sketch.o" "$OUT/core.a" -L"$OUT" -lm
"$BIN/avr-objcopy" -O ihex -R .eeprom "$OUT/sketch.elf" "$OUT/sketch.hex"
"$BIN/avr-size" --format=avr --mcu=$MCU "$OUT/sketch.elf" | sed -n '2,8p'

echo "== upload to $PORT =="
"$BIN/avrdude" -C "$(cygpath -w "$CONF")" -q -q -p$MCU -carduino -P"$PORT" -b115200 -D \
  -Uflash:w:"$(cygpath -w "$OUT/sketch.hex")":i
echo "upload finished"
