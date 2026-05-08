#include <windows.h>

#include <stdio.h>
#include <stdlib.h>

typedef struct gm_time {
    unsigned char sec;
    unsigned char min;
    unsigned char hour;
    unsigned char day;
    unsigned char month;
    unsigned char year;
} gm_time;

typedef struct gm_store {
    gm_time time;
    unsigned char range;
    unsigned char mode;
    unsigned char units;
    float value;
} gm_store;

typedef long(__stdcall *gm0_newgm_fn)(long port, long mode);
typedef long(__stdcall *gm0_startconnect_fn)(long handle);
typedef long(__stdcall *gm0_killgm_fn)(long handle);
typedef int(__stdcall *gm0_getconnect_fn)(long handle);
typedef long(__stdcall *gm0_isnewdata_fn)(long handle);
typedef long(__stdcall *gm0_getrange_fn)(long handle);
typedef long(__stdcall *gm0_getunits_fn)(long handle);
typedef long(__stdcall *gm0_getmode_fn)(long handle);
typedef double(__stdcall *gm0_getvalue_fn)(long handle);
typedef long(__stdcall *gm0_enabledebug_fn)(void);

static FARPROC require_symbol(HMODULE module, const char *name) {
    FARPROC proc = GetProcAddress(module, name);
    if (proc == NULL) {
        fprintf(stderr, "missing symbol %s (error=%lu)\n", name, GetLastError());
        ExitProcess(2);
    }
    return proc;
}

int main(int argc, char **argv) {
    const char *dll_path = argc > 1 ? argv[1] : "e:\\Github\\RAPID\\lib\\gm0.dll";
    long port = argc > 2 ? strtol(argv[2], NULL, 10) : -1;
    long mode = argc > 3 ? strtol(argv[3], NULL, 10) : 1;
    DWORD start_tick = 0;
    HMODULE module = LoadLibraryA(dll_path);
    long handle = -1;
    int connected = 0;
    int saw_new_data = 0;

    if (module == NULL) {
        fprintf(stderr, "LoadLibrary failed for %s (error=%lu)\n", dll_path, GetLastError());
        return 1;
    }

    gm0_enabledebug_fn gm0_enabledebug = (gm0_enabledebug_fn)require_symbol(module, "gm0_enabledebug");
    gm0_newgm_fn gm0_newgm = (gm0_newgm_fn)require_symbol(module, "gm0_newgm");
    gm0_startconnect_fn gm0_startconnect = (gm0_startconnect_fn)require_symbol(module, "gm0_startconnect");
    gm0_killgm_fn gm0_killgm = (gm0_killgm_fn)require_symbol(module, "gm0_killgm");
    gm0_getconnect_fn gm0_getconnect = (gm0_getconnect_fn)require_symbol(module, "gm0_getconnect");
    gm0_isnewdata_fn gm0_isnewdata = (gm0_isnewdata_fn)require_symbol(module, "gm0_isnewdata");
    gm0_getrange_fn gm0_getrange = (gm0_getrange_fn)require_symbol(module, "gm0_getrange");
    gm0_getunits_fn gm0_getunits = (gm0_getunits_fn)require_symbol(module, "gm0_getunits");
    gm0_getmode_fn gm0_getmode = (gm0_getmode_fn)require_symbol(module, "gm0_getmode");
    gm0_getvalue_fn gm0_getvalue = (gm0_getvalue_fn)require_symbol(module, "gm0_getvalue");

    DeleteFileA("c:\\rs232.log");
    gm0_enabledebug();

    handle = gm0_newgm(port, mode);
    printf("handle=%ld port=%ld mode=%ld\n", handle, port, mode);
    fflush(stdout);
    if (handle < 0) {
        fprintf(stderr, "gm0_newgm failed\n");
        FreeLibrary(module);
        return 3;
    }

    printf("startconnect=%ld\n", gm0_startconnect(handle));
    fflush(stdout);

    start_tick = GetTickCount();
    while ((GetTickCount() - start_tick) < 30000) {
        printf("before_getconnect elapsed_ms=%lu\n", GetTickCount() - start_tick);
        fflush(stdout);
        connected = gm0_getconnect(handle);
        printf("after_getconnect connected=%d elapsed_ms=%lu\n", connected, GetTickCount() - start_tick);
        fflush(stdout);
        if (connected) {
            break;
        }
        Sleep(100);
    }
    printf("connected=%d elapsed_ms=%lu\n", connected, GetTickCount() - start_tick);

    if (connected) {
        start_tick = GetTickCount();
        while ((GetTickCount() - start_tick) < 5000) {
            saw_new_data = gm0_isnewdata(handle);
            if (saw_new_data) {
                break;
            }
            Sleep(100);
        }
        printf("isnewdata=%d elapsed_ms=%lu\n", saw_new_data, GetTickCount() - start_tick);
        printf(
            "value=%0.9f mode=%ld units=%ld range=%ld\n",
            gm0_getvalue(handle),
            gm0_getmode(handle),
            gm0_getunits(handle),
            gm0_getrange(handle)
        );
    }

    printf("killgm=%ld\n", gm0_killgm(handle));
    FreeLibrary(module);

    if (GetFileAttributesA("c:\\rs232.log") != INVALID_FILE_ATTRIBUTES) {
        printf("debug_log=c:\\rs232.log\n");
    } else {
        printf("debug_log=missing\n");
    }
    return 0;
}