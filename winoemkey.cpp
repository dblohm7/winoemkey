// cl -EHsc -MT winoemkey.cpp -link -manifest:embed

#include <algorithm>
#include <fstream>
#include <memory>
#include <string>

#include <windows.h>
#include <commctrl.h>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(linker, "")
#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif

using namespace std;

#pragma pack(push, 1)
struct MSDM_HEADER
{
  DWORD   mSignature;
  DWORD   mLengthInBytes;
  BYTE    mRevision;
  BYTE    mChecksum;
  BYTE    mOemId[6];
  BYTE    mOemTableId[8];
  DWORD   mOemTableIdRevision;
  DWORD   mCreatorId;
  DWORD   mCreatorRevision;
};

struct MS_LICENSING_HEADER
{
  DWORD64 mVer1;  // 0x1
  DWORD64 mVer2;  // 0X1
  DWORD   mActivationKeyLen;
/*
  BYTE key[mActivationKeyLen];
  // padding to fix alignment
  ... etc
*/
};
#pragma pack(pop)

const DWORD kMSDMSig = 'MDSM';

void Error(const char* aMsg)
{
  MessageBox(NULL, aMsg, "Error", MB_OK | MB_ICONSTOP);
}

bool ToWide(const string& aStr, wstring& aWStr)
{
  int len = MultiByteToWideChar(CP_UTF8, 0, aStr.c_str(), aStr.size(),
                                nullptr, 0);
  if (!len) {
    return false;
  }
  aWStr.resize(len);
  int finalLen = MultiByteToWideChar(CP_UTF8, 0, aStr.c_str(), aStr.size(),
                                     const_cast<wchar_t*>(aWStr.c_str()),
                                     aWStr.size());
  return finalLen == len;
}

void Info(const string& aMsg, const wstring& aMsg2)
{
  wstring wMsg;
  if (!ToWide(aMsg, wMsg) ||
      FAILED(TaskDialog(NULL, NULL, L"Info", wMsg.c_str(), aMsg2.c_str(),
                        TDCBF_OK_BUTTON, TD_INFORMATION_ICON, nullptr))) {
    MessageBox(NULL, aMsg.c_str(), "Info", MB_OK | MB_ICONASTERISK);
  }
}

bool CopyToClipboard(const string &aText)
{
  auto clipDel = [](void*) { CloseClipboard(); };
  unique_ptr<void, decltype(clipDel)> clip((void*)OpenClipboard(NULL), clipDel);
  if (!OpenClipboard(NULL)) {
    return false;
  }
  if (!EmptyClipboard()) {
    return false;
  }
  auto globDel = [](void* handle) { GlobalFree((HGLOBAL)handle); };
  unique_ptr<void, decltype(globDel)> clipData((void*)GlobalAlloc(GMEM_MOVEABLE, aText.size() + 1), globDel);
  if (!clipData) {
    return false;
  }
  char* pClipData = reinterpret_cast<char*>(GlobalLock((HGLOBAL)clipData.get()));
  if (!pClipData) {
    return false;
  }
  strcpy(pClipData, aText.c_str());
  GlobalUnlock((HGLOBAL)clipData.get());
  pClipData = nullptr;
  if (!SetClipboardData(CF_TEXT, (HGLOBAL)clipData.get())) {
    return false;
  }
  clipData.release(); // The clipboard owns the HGLOBAL now
  return true;
}

int CALLBACK WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
  auto comDel = [](void* hr) { CoUninitialize(); };
  unique_ptr<void, decltype(comDel)>
    com((void*)CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED), comDel);
  UINT reqdSize = EnumSystemFirmwareTables('ACPI', nullptr, 0);
  if (!reqdSize) {
    Error("EnumSystemFirmwareTables failed to return size of ACPI data");
    return 1;
  }
  unsigned int numElements = reqdSize / sizeof(DWORD);
  auto buffer = make_unique<DWORD[]>(numElements);
  UINT result = EnumSystemFirmwareTables('ACPI', buffer.get(), reqdSize);
  if (result != reqdSize) {
    Error("EnumSystemFirmwareTables failed");
    return 1;
  }
  if (find(&buffer[0], &buffer[numElements], kMSDMSig) == &buffer[numElements]) {
    Error("Microsoft Licensing table not found");
    return 1;
  }
  reqdSize = GetSystemFirmwareTable('ACPI', kMSDMSig, nullptr, 0);
  if (!reqdSize) {
    Error("GetSystemFirmwareTable failed to return size of table");
    return 1;
  }
  auto table = make_unique<char[]>(reqdSize);
  result = GetSystemFirmwareTable('ACPI', kMSDMSig, table.get(), reqdSize);
  if (result != reqdSize) {
    Error("GetSystemFirmwareTable failed");
    return 1;
  }
  MS_LICENSING_HEADER* licHeader = reinterpret_cast<MS_LICENSING_HEADER*>(table.get() + sizeof(MSDM_HEADER));
  string key(table.get() + sizeof(MSDM_HEADER) + sizeof(MS_LICENSING_HEADER),
             licHeader->mActivationKeyLen);
  bool copiedToClipboard = false;
  if (CopyToClipboard(key)) {
    copiedToClipboard = true;
  }
  Info(key.c_str(), copiedToClipboard ? L"This data has been copied to the clipboard." : L"Could not copy data to clipboard.");
  return 0;
}

