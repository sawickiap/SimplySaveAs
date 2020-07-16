#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include <cstdlib>
#include <cassert>
#include <cmath>

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR cmdLine, int cmdShow)
{
    MessageBox(NULL, L"Hello World!", L"Title", MB_OK);
}
