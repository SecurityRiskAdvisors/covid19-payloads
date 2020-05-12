#define EXPORTING_DLL
#include "ta505.h"
#include <windows.h>
#pragma comment(lib,"User32.lib")

BOOL APIENTRY DllMain( HANDLE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved )
{
    return TRUE;
}

void main()
{
    MessageBox( NULL, TEXT("hello world"), TEXT("hello world"), MB_OK);
}