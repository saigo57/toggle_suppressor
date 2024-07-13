#include <windows.h>
#include <mmdeviceapi.h>
#include <audiopolicy.h>
#include <iostream>
#include <memory>
#include <tlhelp32.h>

std::wstring PrintProcessNameAndID(DWORD proccessId)
{
	HANDLE hToolhelp32Snapshot = CreateToolhelp32Snapshot(
		TH32CS_SNAPPROCESS,
		NULL
	);

	if (hToolhelp32Snapshot == INVALID_HANDLE_VALUE) {
		return std::wstring(L"");
	}

	PROCESSENTRY32 pe32;
	pe32.dwSize = sizeof(PROCESSENTRY32);

	if (Process32First(hToolhelp32Snapshot, &pe32)) {
		do {
			if (proccessId == pe32.th32ProcessID) {
				return std::wstring(pe32.szExeFile);
			}
		} while (Process32Next(hToolhelp32Snapshot, &pe32));
	} else {
		CloseHandle(hToolhelp32Snapshot);
	}

	return std::wstring(L"");
}

std::wstring charTowstring(const char* orig)
{
	// 確保するワイド文字用バッファのサイズは、バイト数ではなく文字数を指定する。
	size_t newsize = strlen(orig) + 1;
	wchar_t* wc = new wchar_t[newsize];

	// 変換.
	size_t convertedChars = 0;
	mbstowcs_s(&convertedChars, wc, newsize, orig, _TRUNCATE);

	std::wstring ret(wc);
	return ret;
}

int main(int argc, char* argv[]) {
	if (argc != 2) {
		std::cout << "param is invalid." << std::endl;
		return -1;
	}

	std::wstring targetExeName = charTowstring(argv[1]);

	if ( CoInitializeEx(0, COINIT_MULTITHREADED) != S_OK ) {
		std::cout << "Failed to initialize COM" << std::endl;
		return -1;
	}

	IMMDeviceEnumerator* pEnumerator;
	if ( CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, IID_PPV_ARGS(&pEnumerator)) != S_OK ) {
		std::cout << "Failed to create device enumerator" << std::endl;
		return -1;
	}

	IMMDevice* defaultDevice = nullptr;
	if (pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &defaultDevice) != S_OK) {
		std::cout << "Failed to get default audio endpoint" << std::endl;
		return -1;
	}

	IAudioSessionManager2* audioSessionManager = nullptr;
	defaultDevice->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL, nullptr, (void**)&audioSessionManager);

	IAudioSessionEnumerator* sessionEnumerator = nullptr;
	audioSessionManager->GetSessionEnumerator(&sessionEnumerator);

	int sessionCount = 0;
	sessionEnumerator->GetCount(&sessionCount);

	std::cout << "sessionCount: " << sessionCount << std::endl;

	for (int i = 0; i < sessionCount; i++) {
		IAudioSessionControl* sessionControl = nullptr;
		sessionEnumerator->GetSession(i, &sessionControl);

		IAudioSessionControl2* sessionControl2 = nullptr;
		sessionControl->QueryInterface(__uuidof(IAudioSessionControl2), (void**)&sessionControl2);

		DWORD processId;
		sessionControl2->GetProcessId(&processId);
		
		std::wstring exeName = PrintProcessNameAndID(processId);

		if ( exeName.find(targetExeName) != std::wstring::npos ) {
			ISimpleAudioVolume* audioVolume = nullptr;
			sessionControl->QueryInterface(__uuidof(ISimpleAudioVolume), (void**)&audioVolume);

			BOOL mute = FALSE;
			audioVolume->GetMute(&mute);
			audioVolume->SetMute(!mute, nullptr);
			audioVolume->Release();
		}

		sessionControl2->Release();
		sessionControl->Release();
	}

	sessionEnumerator->Release();
	audioSessionManager->Release();
	defaultDevice->Release();
	pEnumerator->Release();
	CoUninitialize();

	return 0;
}
