#define UNICODE

#pragma comment(lib,"rpcrt4")
#pragma comment(lib,"shlwapi")
#pragma comment(linker,"/manifestdependency:\"type='win32' name = 'Microsoft.Windows.Common-Controls' version = '6.0.0.0' processorArchitecture = '*' publicKeyToken = '6595b64144ccf1df' language = '*'\"")

#import "C:\Program Files (x86)\Common Files\Microsoft Shared\DAO\dao360.dll" rename_namespace("DAO") rename("EOF", "adoEOF")
#import "C:\Program Files (x86)\Common Files\System\ado\msado60.tlb" no_namespace rename("EOF", "EndOfFile")

#include <windows.h>
#include <shlwapi.h>
#include <odbcinst.h>

TCHAR szClassName[] = TEXT("CreateDatabase");

BOOL CreateDatabase(HWND hWnd, LPCTSTR lpszFilePath)
{
	CoInitialize(NULL);
	TCHAR szAttributes[1024];
	wsprintf(szAttributes, TEXT("CREATE_DB=\"%s\" General\0"), lpszFilePath);
	if (!SQLConfigDataSource(hWnd, ODBC_ADD_DSN, TEXT("Microsoft Access Driver (*.mdb)"), szAttributes))
	{
		CoUninitialize();
		return FALSE;
	}
	CoUninitialize();
	return TRUE;
}

BOOL SQLExecute(HWND hWnd, LPCTSTR lpszMDBFilePath, LPCTSTR lpszSQL)
{
	HRESULT hr;
	_ConnectionPtr pCon(NULL);
	hr = pCon.CreateInstance(__uuidof(Connection));
	if (FAILED(hr))
	{
		return FALSE;
	}
	TCHAR szString[1024];
	wsprintf(szString, TEXT("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;"), lpszMDBFilePath);
	hr = pCon->Open(szString, _bstr_t(""), _bstr_t(""), adOpenUnspecified);
	if (FAILED(hr))
	{
		return FALSE;
	}
	BOOL bRet = TRUE;
	try
	{
		_CommandPtr pCommand(NULL);
		pCommand.CreateInstance(__uuidof(Command));
		pCommand->ActiveConnection = pCon;
		pCommand->CommandText = lpszSQL;
		pCommand->Execute(NULL, NULL, adCmdText);
	}
	catch (_com_error&e)
	{
		MessageBox(hWnd, e.Description(), 0, 0);
		bRet = FALSE;
	}
	pCon->Close();
	pCon = NULL;
	return TRUE;
}

BOOL CreateGUID(TCHAR *szGUID)
{
	GUID m_guid = GUID_NULL;
	HRESULT hr = UuidCreate(&m_guid);
	if (HRESULT_CODE(hr) != RPC_S_OK){ return FALSE; }
	if (m_guid == GUID_NULL){ return FALSE; }
	wsprintf(szGUID, TEXT("{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}"),
		m_guid.Data1, m_guid.Data2, m_guid.Data3,
		m_guid.Data4[0], m_guid.Data4[1], m_guid.Data4[2], m_guid.Data4[3],
		m_guid.Data4[4], m_guid.Data4[5], m_guid.Data4[6], m_guid.Data4[7]);
	return TRUE;
}

BOOL CreateTempDirectory(LPTSTR pszDir)
{
	TCHAR szGUID[39];
	if (GetTempPath(MAX_PATH, pszDir) == 0)return FALSE;
	if (CreateGUID(szGUID) == 0)return FALSE;
	if (PathAppend(pszDir, szGUID) == 0)return FALSE;
	if (CreateDirectory(pszDir, 0) == 0)return FALSE;
	return TRUE;
}

BOOL CompactDatabase(HWND hWnd, LPCTSTR lpszMDBFilePath, LPCTSTR lpszPassword = 0)
{
	BOOL bRet = FALSE;
	DAO::_DBEngine* pEngine = NULL;
	HRESULT hr = CoCreateInstance(__uuidof(DAO::DBEngine), NULL, CLSCTX_ALL, IID_IDispatch, (LPVOID*)&pEngine);
	if (SUCCEEDED(hr) && pEngine)
	{
		hr = -1;
		TCHAR szTempDirectoryPath[MAX_PATH];
		if (CreateTempDirectory(szTempDirectoryPath))
		{
			PathAppend(szTempDirectoryPath, TEXT("TmpDatabase.mdb"));
			try
			{
				TCHAR szString[1024];
				if (lpszPassword)
				{
					wsprintf(szString, TEXT(";pwd=%s"), lpszPassword);
					pEngine->CompactDatabase((_bstr_t)lpszMDBFilePath, (_bstr_t)szTempDirectoryPath, vtMissing, vtMissing, szString);
				}
				else
				{
					pEngine->CompactDatabase((_bstr_t)lpszMDBFilePath, (_bstr_t)szTempDirectoryPath);
				}
			}
			catch (_com_error& e)
			{
				MessageBox(hWnd, e.Description(), 0, 0);
			}
			if (SUCCEEDED(hr))
			{
				if (MoveFileEx(szTempDirectoryPath, lpszMDBFilePath, MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
				{
					bRet = TRUE;
				}
			}
		}
		pEngine->Release();
		pEngine = NULL;
	}
	return bRet;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hButton;
	switch (msg)
	{
	case WM_CREATE:
		CreateWindow(TEXT("BUTTON"), TEXT("データベースを作成する"), WS_VISIBLE | WS_CHILD, 10, 10, 256, 32, hWnd, (HMENU)100, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == 100)
		{
			EnableWindow(hButton, FALSE);
			OPENFILENAME ofn = { sizeof(OPENFILENAME) };
			TCHAR szFileName[MAX_PATH] = TEXT("sample.mdb");
			ofn.hwndOwner = hWnd;
			ofn.lpstrFilter = TEXT("Access Database Files (*.mdb)\0*.mdb\0All Files (*.*)\0*.*\0");
			ofn.lpstrFile = szFileName;
			ofn.nMaxFile = MAX_PATH;
			ofn.lpstrDefExt = TEXT("mdb");
			ofn.Flags = OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
			if (GetSaveFileName(&ofn))
			{
				if (CreateDatabase(hWnd, szFileName))
				{
					SQLExecute(hWnd, szFileName, TEXT("CREATE TABLE 名簿(名前 VARCHAR (255), 特技 VARCHAR (255));"));
					SQLExecute(hWnd, szFileName, TEXT("INSERT INTO 名簿(名前,特技)VALUES('山田太郎','スイカ割り');"));
					SQLExecute(hWnd, szFileName, TEXT("INSERT INTO 名簿(名前,特技)VALUES('山田花子','早口言葉');"));
					CompactDatabase(hWnd, szFileName);
				}
			}
			EnableWindow(hButton, TRUE);
		}
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0, IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("Access データベースファイルを作成する"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
		);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return msg.wParam;
}
