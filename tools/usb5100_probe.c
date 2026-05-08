#include <windows.h>

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

typedef unsigned int(__stdcall *openUSB5100_fn)(void);
typedef void(__stdcall *closeUSB5100_fn)(unsigned int handle);
typedef int(__stdcall *scpiCommand_fn)(unsigned int handle, const char *command, char *response, int length);

static void print_last_error(const char *prefix, DWORD error_code) {
    char message[1024];
    DWORD flags = FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
    DWORD length = FormatMessageA(
        flags,
        NULL,
        error_code,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        message,
        (DWORD)sizeof(message),
        NULL
    );

    if (length == 0) {
        fprintf(stderr, "%s error=%lu\n", prefix, error_code);
        return;
    }

    while (length > 0 && (message[length - 1] == '\r' || message[length - 1] == '\n')) {
        message[length - 1] = '\0';
        length -= 1;
    }
    fprintf(stderr, "%s error=%lu message=%s\n", prefix, error_code, message);
}

static FARPROC require_symbol(HMODULE module, const char *name) {
    FARPROC proc = GetProcAddress(module, name);
    if (proc == NULL) {
        print_last_error(name, GetLastError());
        ExitProcess(2);
    }
    return proc;
}

static void print_search_result(const char *name) {
    char resolved[MAX_PATH];
    DWORD length = SearchPathA(NULL, name, NULL, MAX_PATH, resolved, NULL);
    if (length == 0 || length >= MAX_PATH) {
        printf("search %s=missing\n", name);
        return;
    }
    printf("search %s=%s\n", name, resolved);
}

static int resolve_default_dll_path(char *resolved, DWORD resolved_length) {
    DWORD length = SearchPathA(NULL, "usb5100.dll", NULL, resolved_length, resolved, NULL);
    if (length == 0 || length >= resolved_length) {
        return 0;
    }
    return 1;
}

static int run_scpi_command(
    scpiCommand_fn scpiCommand,
    unsigned int handle,
    const char *command,
    char *response,
    int response_length
) {
    size_t command_length = strlen(command) + 1;
    char *command_buffer = (char *)malloc(command_length);
    int result = 0;

    if (command_buffer == NULL) {
        fprintf(stderr, "failed to allocate command buffer\n");
        return -1;
    }

    memcpy(command_buffer, command, command_length);
    ZeroMemory(response, response_length);
    result = scpiCommand(handle, command_buffer, response, response_length);
    free(command_buffer);
    return result;
}

static int run_legacy_probe(
    const char *dll_path,
    openUSB5100_fn openUSB5100,
    closeUSB5100_fn closeUSB5100,
    scpiCommand_fn scpiCommand
) {
    char response[80];
    unsigned int handle = 0;
    int scpi_status = 0;

    handle = openUSB5100();
    printf("openUSB5100 handle=%u\n", handle);
    fflush(stdout);
    if (handle == 0) {
        return 3;
    }

    scpi_status = run_scpi_command(scpiCommand, handle, "*IDN?", response, (int)sizeof(response));
    printf("scpi *IDN? status=%d response=%s\n", scpi_status, response);

    scpi_status = run_scpi_command(scpiCommand, handle, ":MEASURE:FLUX?", response, (int)sizeof(response));
    printf("scpi :MEASURE:FLUX? status=%d response=%s\n", scpi_status, response);

    closeUSB5100(handle);
    printf("closeUSB5100 handle=%u\n", handle);
    printf("loaded=%s\n", dll_path);
    return 0;
}

static int run_cli_command(
    const char *dll_path,
    const char *command_name,
    const char *command_arg,
    openUSB5100_fn openUSB5100,
    closeUSB5100_fn closeUSB5100,
    scpiCommand_fn scpiCommand
) {
    char response[80];
    unsigned int handle = 0;
    int scpi_status = 0;
    const char *scpi_text = NULL;

    handle = openUSB5100();
    if (handle == 0) {
        fprintf(stderr, "openUSB5100 failed\n");
        return 3;
    }

    if (strcmp(command_name, "status") == 0) {
        scpi_text = "*IDN?";
    } else if (strcmp(command_name, "read") == 0) {
        scpi_text = ":MEASURE:FLUX?";
    } else if (strcmp(command_name, "command") == 0) {
        if (command_arg == NULL || command_arg[0] == '\0') {
            fprintf(stderr, "command mode requires a SCPI command argument\n");
            closeUSB5100(handle);
            return 4;
        }
        scpi_text = command_arg;
    } else {
        fprintf(stderr, "unknown command: %s\n", command_name);
        closeUSB5100(handle);
        return 4;
    }

    scpi_status = run_scpi_command(scpiCommand, handle, scpi_text, response, (int)sizeof(response));
    if (scpi_status != 0) {
        fprintf(stderr, "scpi command failed status=%d command=%s response=%s\n", scpi_status, scpi_text, response);
        closeUSB5100(handle);
        return 5;
    }

    printf("status=ok\n");
    printf("dll=%s\n", dll_path);
    printf("command=%s\n", scpi_text);
    printf("response=%s\n", response);

    closeUSB5100(handle);
    return 0;
}

int main(int argc, char **argv) {
    const char *dll_path = NULL;
    const char *command_name = NULL;
    const char *command_arg = NULL;
    int arg_index = 1;
    char dll_dir[MAX_PATH];
    char resolved_dll_path[MAX_PATH];
    HMODULE module = NULL;
    int exit_code = 0;

    openUSB5100_fn openUSB5100 = NULL;
    closeUSB5100_fn closeUSB5100 = NULL;
    scpiCommand_fn scpiCommand = NULL;

    while (arg_index < argc) {
        if (strcmp(argv[arg_index], "--dll") == 0) {
            if ((arg_index + 1) >= argc) {
                fprintf(stderr, "--dll requires a path\n");
                return 4;
            }
            dll_path = argv[arg_index + 1];
            arg_index += 2;
            continue;
        }

        command_name = argv[arg_index];
        if ((arg_index + 1) < argc) {
            command_arg = argv[arg_index + 1];
        }
        break;
    }

    if (dll_path == NULL) {
        if (!resolve_default_dll_path(resolved_dll_path, MAX_PATH)) {
            fprintf(stderr, "usb5100.dll was not found. Pass --dll <path> or add the vendor DLL directory to PATH.\n");
            return 4;
        }
        dll_path = resolved_dll_path;
    }

    if (GetFullPathNameA(dll_path, MAX_PATH, dll_dir, NULL) == 0) {
        strncpy_s(dll_dir, MAX_PATH, dll_path, _TRUNCATE);
    }

    {
        char *last_sep = strrchr(dll_dir, '\\');
        if (last_sep != NULL) {
            *last_sep = '\0';
            SetDllDirectoryA(dll_dir);
            printf("dll_dir=%s\n", dll_dir);
        }
    }

    print_search_result("usb5100.dll");
    print_search_result("libusb0.dll");

    module = LoadLibraryA(dll_path);
    if (module == NULL) {
        print_last_error("LoadLibrary usb5100.dll", GetLastError());
        return 1;
    }
    printf("loaded=%s\n", dll_path);

    openUSB5100 = (openUSB5100_fn)require_symbol(module, "openUSB5100");
    closeUSB5100 = (closeUSB5100_fn)require_symbol(module, "closeUSB5100");
    scpiCommand = (scpiCommand_fn)require_symbol(module, "scpiCommand");

    if (command_name == NULL) {
        exit_code = run_legacy_probe(dll_path, openUSB5100, closeUSB5100, scpiCommand);
    } else {
        exit_code = run_cli_command(dll_path, command_name, command_arg, openUSB5100, closeUSB5100, scpiCommand);
    }

    FreeLibrary(module);
    return exit_code;
}