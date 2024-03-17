//
//		SpoutWallPaper
//
//		A system tray application for Spout Senders or video files
//		as live desktop wallpaper or show the Bing daily wallpaper.
//
//		Spout receiver
//		  If a Spout sender is running, it will be displayed immediately.
//		  If not running, a sender will be detected as soon as it starts
//		  For images or video, the senders are not displayed but can be selected.
//		  To select a sender :
//		    o Find SpoutWallPaper in the TaskBar tray area
//		    o Right mouse click to select the sender
//
//		Video player
//		  Select "Video" from the menu and choose the video file
//
//		Image
//		  Select "Image" from the menu and choose the image file
//
//		Daily wallpaper
//		  Select "Daily" from the menu.
//
//		Slideshow
//		  Select image folder, enter slide duration and check "Random" if required
//
//		"About" for details.
//
//		At program close there is an option to keep the new wallpaper or restore the original.
//
//		Based on 2013 Spout 2.006 "SpoutTray" which was based on :
//		http://www.codeproject.com/Articles/4768/Basic-use-of-Shell_NotifyIcon-in-Win32
//		and video image capture using FFmpeg based on :
//		https://github.com/leadedge/SpoutDXvideo
//
//		Wallpaper code based on : 
//		https://github.com/ohkashi/LiveWallpaper
//		Daily wallpaper based on :
//		https://github.com/aisyshk/DailyWallpapers
//		Others :
//		https://github.com/search?q=video+wallpaper+language%3AC%2B%2B&type=repositories
//		https://github.com/search?q=live+wallpaper+language%3AC%2B%2B&type=repositories
//		And many others :
//		https://github.com/search?q=wallpaper+language%3AC%2B%2B&type=repositories
//
//		Bing wallpaper can also be downloaded daily using a Microsoft app :
//		https://www.microsoft.com/en-us/bing/bing-wallpaper
//		Image archive :
//		https://bing.wallpaper.pics/
//
// =========================================================================
//
//               Copyright(C) 2024 Lynn Jarvis.
//               https://www.spout.zeal.co
//
// This program is free software : you can redistribute it and/or modify
// it under the terms of the GNU Lesser General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
// GNU Lesser General Public License for more details.
// 
// You should have received a copy of the GNU Lesser General Public License
// along with this program.If not, see <http://www.gnu.org/licenses/>.
//
// =========================================================================
//
//		24-02-24 - Create project
//		25.01.24 - Add option for video files
//		27.01.24 - Add video selection dialog
//		28.02.24 - Add daily wallpaper
//		29.02.24 - Add image selection dialog
//		02.02.24 - Add daily image description for about box and exit
//		03.02.24 - Version 1.001
//		06.02.24 - Use 'find' instead or 'rfind' for image url
//				   Remove '?' from daily image title
//				   Introduce bCurrentWallpaper to avoid repeated set in RestoreWallPaper
//				   Render() - when the sender closes, restore daily or original wallpaper
//				   Version 1.002
//		10.03.24 - Introduce bCurrentWallpaper if showing original wallpaper for video
//				   SPIF_SENDCHANGE instead of SPIF_UPDATEINIFILE for wallpaper change
//				   unless restoring or setting wallpaper
//				 - Add slideshow option with image path, slide duration and random change
//		12.03.24 - Update slide image details for About and Exit
//		17.03.24 - Version 1.003
//

#include "stdafx.h"
#include <windows.h>
#include <stdio.h> // for printf if used
#include <shlobj.h>
#include <commdlg.h> // for explorer dialog
#include "..\..\SpoutDirectX\SpoutDX\SpoutDX.h"
#include "resource.h"

// for PathStripPath
#include <Shlwapi.h>
#pragma comment (lib, "Shlwapi.lib")

// For UrlDownloadToFile
#include <Urlmon.h>
#pragma comment (lib, "Urlmon.lib")

#define TRAYICONID	1        // ID number for the Notify Icon
#define SWM_TRAYMSG	WM_APP   // The message ID sent to our window
#define SWM_EXIT WM_APP + 13 // Close the window
#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst = NULL;  // current instance
NOTIFYICONDATA niData{}; // notify icon data
HWND hWndMain = NULL;
HANDLE g_hMutex = NULL;                // Application mutex
WCHAR szTitle[MAX_LOADSTRING]{};       // title bar text
WCHAR szWindowClass[MAX_LOADSTRING]{}; // main window class name

// Variables to enumerate Spout senders
spoutDX spout;
SharedTextureInfo info;
std::set<std::string> Senders;
std::set<std::string>::iterator iter;
std::string namestring;
char name[512]{};

// Spout receiver
spoutDX receiver;                       // Receiver object
HWND g_hWnd = NULL;                     // Window handle
unsigned char* g_pixelBuffer = nullptr; // Receiving pixel buffer
unsigned int g_SenderWidth = 0;         // Received sender width
unsigned int g_SenderHeight = 0;        // Received sender height
HWND g_WorkerHwnd = NULL;               // Worker window handle
void Render();

// For the Bing daily wallpaper image
std::string g_wallpaperpath;      // Current wallpaper image
bool bCurrentWallpaper = false;   // Showing current wallpaper
std::string g_dailywallpaperpath; // Daily wallpaper image
bool bDailyWallpaper = false;     // Downloaded daily wallpaper
bool bShowDaily = false;          // Showing daily wallpaper
std::string copyright;            // Description for about and exit

// For slideshow
bool bSlideShow = false;
bool OpenFolder(char* filepath, int maxchars);
char g_slideshowpath[MAX_PATH]{}; // Slideshow folder
char startingfolder[MAX_PATH]{}; // Initial browse dialog folder
int nCurrentImage = 0;
DWORD g_slideshowtime = 30; // Seconds per frame
double g_start = 0.0; // Start time
std::vector<std::string> slidenames; // Slideshow file names
int GetImageFiles(const char* spath, std::vector<std::string>& filenames);

// For FFmpeg video player
std::string g_videopath;            // The full video path
unsigned char g_SenderName[256]={}; // Sender name
float g_FrameRate = 30.0f;          // Video frame rate
double g_SenderFps = 0.0;           // For fps display averaging
std::string g_exePath;              // Executable location
std::string g_ffmpegPath;           // FFmpeg location
std::string g_input;                // Input string to FFmpeg
FILE* g_pipein = nullptr;           // Pipe for FFmpeg

// Forward declarations
BOOL InitInstance(HINSTANCE, int);
BOOL OnInitDialog(HWND hWnd);
void ShowContextMenu(HWND hWnd);
ULONGLONG GetDllVersion(LPCTSTR lpszDllName);
void RestoreWallPaper();
bool OpenVideo(std::string filePath);
void CloseVideo();
bool OpenFile(char* filepath, int maxchars, bool bVideo = false);
bool ffprobe(std::string videoPath);
INT_PTR CALLBACK DlgProc(HWND, UINT, WPARAM, LPARAM);

// Slide duration
bool SelectSlideDuration();
static INT_PTR CALLBACK SlideDuration(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
static int durationindex = 0; // for the combobox
static int duration[13]= { 2, 4, 10, 30, 60, 120, 240, 600, 1800, 3600, 7200, 14400, 36000 }; // seconds
static bool bRandom = false;

// For folder selection
INT CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM lp, LPARAM pData);
BOOL CALLBACK EnumWindowsProc(HWND hWnd, LPARAM lParam);

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	MSG msg{};
	HACCEL hAccelTable = NULL;

	// Console window so printf works
	FILE* pCout;
	AllocConsole();
	freopen_s(&pCout, "CONOUT$", "w", stdout); 
	printf("SpoutWallPaper\n");

	// Try to open the application mutex
	g_hMutex = OpenMutexA(MUTEX_ALL_ACCESS, 0, "SpoutWallpaper");
	// If it exists, don't open again
	if (g_hMutex) {
		SpoutMessageBox("SpoutWallpaper is already running");
		// Close it now, otherwise it is never released
		CloseHandle(g_hMutex);
		return 0;
	}

	// Executable location
	char exePath[MAX_PATH]={};
	GetModuleFileNameA(NULL, exePath, MAX_PATH); // Full path of the executable
	PathRemoveFileSpecA(exePath);
	g_exePath = exePath;

	// Look for FFmpeg
	g_ffmpegPath = g_exePath;
	g_ffmpegPath += "\\DATA\\FFMPEG\\ffmpeg.exe";
	// Look for FFprobe.exe
	std::string ffpath = g_exePath;
	ffpath += "\\DATA\\FFMPEG\\ffprobe.exe";
	// Location should be : \DATA\FFMPEG
	if (_access(g_ffmpegPath.c_str(), 0) == -1 || _access(ffpath.c_str(), 0) == -1) {
		std::string str = "o Go to <a href=\"https://github.com/GyanD/codexffmpeg/releases/\">https://github.com/GyanD/codexffmpeg/releases/</a>\n";
		str += "o Choose the \"Essentials\" build. e.g.ffmpeg-5.1.2-essentials_build.zip\n";
		str += "o Download and unzip the archive\n";
		str += "o Copy bin\\FFmpeg.exe and bin\\FFprobe.exe to : DATA\\FFMPEG";
		SpoutMessageBox("FFmpeg not found", str.c_str(), MB_OK | MB_TOPMOST);
		if (g_hMutex) ReleaseMutex(g_hMutex);
		return 0;
	}

	// Get the current wallpaper to restore if changed
	char path[MAX_PATH]{};
	if (SystemParametersInfoA(SPI_GETDESKWALLPAPER, MAX_PATH, (void*)path, 0)) {
		g_wallpaperpath = path;
		bCurrentWallpaper = true;
	}

	// Get the last slideshow path and slide time
	ReadPathFromRegistry(HKEY_CURRENT_USER, "Software\\Leading Edge\\SpoutWallpaper", "slideshowfolder", g_slideshowpath);
	ReadDwordFromRegistry(HKEY_CURRENT_USER, "Software\\Leading Edge\\SpoutWallpaper", "slideshowtime", &g_slideshowtime);


	//
	// Optional command line : SpoutWallPaper "video name"
	//
	//  Video name can be :
	//     o Full path to the file
	//	   o A file in the executable folder
	//     o A file in the "\DATA\Videos\" folder
	//
	// If no command line, a Spout sender is detected if running
	//
	if (lpCmdLine && *lpCmdLine) {
		std::string videoname = lpCmdLine;
		// Remove double quotes
		videoname.erase(std::remove(videoname.begin(), videoname.end(), '"'), videoname.end());
		// Remove trailing spaces
		size_t pos = videoname.find_last_not_of("\n");
		videoname = (pos == std::string::npos) ? "" : videoname.substr(0, pos + 1);
		// Remove preceding backslash
		pos = videoname.find_first_of("\\");
		if (pos != std::string::npos && pos == 0) {
			videoname = videoname.substr(pos+1);
		}

		if (PathFileExistsA(videoname.c_str())) {
			// Full fath found
			g_videopath = videoname;
		}
		else {
			// Try the executable folder
			g_videopath = g_exePath;
			g_videopath += "\\";
			g_videopath += videoname;
			if (!PathFileExistsA(g_videopath.c_str())) {
				// Try the DATA\Video folder
				g_videopath = g_exePath;
				g_videopath += "\\DATA\\Videos\\";
				g_videopath += videoname;
				if (!PathFileExistsA(g_videopath.c_str())) {
					// The path does not exist
					g_videopath.clear();
				}
			}
		}
	}

	// Initialize global strings
	LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadStringW(hInstance, IDC_LIVE_WALLPAPER, szWindowClass, MAX_LOADSTRING);

	// Find the existing worker window and close it
	HWND hWorker = NULL;
	EnumWindows(EnumWindowsProc, (LPARAM)&hWorker);
	if (hWorker) {
		HWND hWnd = FindWindowEx(hWorker, NULL, szWindowClass, NULL);
		if (hWnd) {
			PostMessage(hWnd, WM_CLOSE, 0, 0);
		}
	}

	// Find the Program Manager window
	HWND progman = FindWindow(L"Progman", NULL);

	// Send message 0x052C to Progman which directs it to spawn a WorkerW 
	// behind the desktop icons. The message is ignored if the window already exists.
	SendMessageTimeout(progman, 0x052C, NULL, NULL, SMTO_NORMAL, 1000, NULL);

	// Enumerate all Windows, until one has the SHELLDLL_DefView as a child. 
	// If found, take its next sibling and assign it to workerw.
	hWorker = NULL;
	EnumWindows(EnumWindowsProc, (LPARAM)&hWorker);
	if (!hWorker) {
		RestoreWallPaper();
		if (g_hMutex) ReleaseMutex(g_hMutex);
		return 0;
	}

	// Set the global worker window handle for drawing
	g_WorkerHwnd = hWorker;

	// Application initialization
	if (!InitInstance(hInstance, nCmdShow)) {
		if (g_hMutex) ReleaseMutex(g_hMutex);
		return 0;
	}

	hAccelTable = LoadAccelerators(hInstance, (LPCTSTR)IDC_STEALTHDIALOG);

	// Set a timer to activate drawing
	// High as possible but low enough to prevent hesitation
	// and less than render rate (at 30fps framerate = 33 msec)
	SetTimer(hWndMain, 1, 30, NULL);

	// Initialize DirectX
	// A device is created in the SpoutDX class.
	if (!receiver.OpenDirectX11()) {
		RestoreWallPaper();
		if (g_hMutex) ReleaseMutex(g_hMutex);
		return 0;
	}

	HANDLE hProcess = GetCurrentProcess();
	
	//
	// Reduce CPU load of the process
	//

	// Set low affinity
	//  0  1  2  3  4  5  6  7  core
	// 01 02 04 08 10 20 40 80  mask bit
	SetProcessAffinityMask(hProcess, 0x03); // Use 2 cores or video load is too slow

	// Set low priority
	SetPriorityClass(hProcess, IDLE_PRIORITY_CLASS);

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0)) {
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg)||
			!IsDialogMessage(msg.hwnd,&msg) ) 
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		if(msg.wParam != WM_QUIT)
			Render();

		Sleep(1); // Reduces CPU load

	}

	KillTimer(hWndMain, 1);

	// Release FFmpeg resources and release buffers
	CloseVideo();

	// Release the receiver
	receiver.ReleaseReceiver();

	// Restore the starting wallpaper unless it was changed and should be retained
	if (bDailyWallpaper && !g_wallpaperpath.empty()) {
		HICON hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_STEALTHDLG));
		SpoutMessageBoxIcon(hIcon);
		std::string confirm = " ";
		if (!copyright.empty()) confirm += copyright;
		if(SpoutMessageBox(NULL, confirm.c_str(), " ", MB_USERICON | MB_YESNO, "Keep new wallpaper ?") == IDYES) {
			g_wallpaperpath.clear();
		}
	}

	RestoreWallPaper();

	// Release DirectX 11 resources
	receiver.CloseDirectX11();

	// Release the application mutex so another instance can be opened
	if (g_hMutex) ReleaseMutex(g_hMutex);

	return (int) msg.wParam;
}

//
// Spout receiver or video from FFmpeg
//
void Render()
{
	// No rendering for wallpaper image
	if (bShowDaily)
		return;
	
	if (!slidenames.empty() && g_start > 0.0) {

		//
		// Slideshow
		//

		// Not showing original wallpaper
		bCurrentWallpaper = false;

		// Calculate frame time for slide duration
		double msecs = ElapsedMicroseconds()/1000.0;
		double elapsed = msecs-g_start; // msec elapsed

		std::string slidepath = g_slideshowpath;
		slidepath += "\\";
		if (nCurrentImage == 0 || elapsed > (double)(g_slideshowtime*1000)) { // seconds to msec
			
			g_start = msecs;
			slidepath += slidenames[nCurrentImage];
			SystemParametersInfoA(SPI_SETDESKWALLPAPER, 0, (void*)slidepath.c_str(), SPIF_SENDCHANGE);
			
			// Not showing original wallpaper
			bCurrentWallpaper = false;

			// Update details for About and Exit
			char filepath[MAX_PATH]{};
			strcpy_s(filepath, slidepath.c_str());
			PathStripPathA(filepath);
			copyright = filepath;
			
			// Update image index
			if (bRandom) {
				nCurrentImage = rand()%(slidenames.size());
			}
			else {
				nCurrentImage++;
				if (nCurrentImage > slidenames.size()-1)
					nCurrentImage = 0;
			}

		}
		return;
	}

	if (g_videopath.empty()) {

		//
		// Spout wallpaper
		//

		// Get pixels from the sender shared texture.
		// ReceiveImage handles sender detection, creation and update.
		// Windows bitmaps are bottom-up. Although the rgba pixel buffer can be flipped by
		// the ReceiveImage function, it can also be drawn upside down as shown further below.
		// Approx 5-8 msec
		if (receiver.ReceiveImage(g_pixelBuffer, g_SenderWidth, g_SenderHeight)) { // RGB = false, invert = false
			
			// IsUpdated() returns true if the sender has changed
			if (receiver.IsUpdated()) {

				// Update globals
				g_SenderWidth = receiver.GetSenderWidth();
				g_SenderHeight = receiver.GetSenderHeight();

				// Update the receiving buffer
				if (g_pixelBuffer) delete[] g_pixelBuffer;
				g_pixelBuffer = new unsigned char[g_SenderWidth * g_SenderHeight * 4];

				// Not showing current wallpaper
				bCurrentWallpaper = false;

				return; // return for next receive
			}
		}
		else {
			// Because SetReceiverName is used to select the sender,
			// the receiver waits for that sender to re-open.
			// Here we need to test to find if it was closed.
			if (!receiver.sendernames.hasSharedInfo(receiver.GetSenderName()))
				// Clear the receiver name
				receiver.SetReceiverName();

			if (g_pixelBuffer) {
				receiver.ReleaseReceiver();
				delete[] g_pixelBuffer;
				g_pixelBuffer = nullptr;
			}

			// If a daily wallpaper has been downloaded, show it
			if (bDailyWallpaper && !g_wallpaperpath.empty()) {
				SystemParametersInfoA(SPI_SETDESKWALLPAPER, 0, (void*)g_dailywallpaperpath.c_str(), SPIF_SENDCHANGE);
				bShowDaily = true;
			}
			else {
				// Otherwise restore wallpaper
				RestoreWallPaper();
			}
		} // endif ReceiveImage
	}
	else {

		//
		// Video using FFmpeg
		//

		// initialize FFmpeg pipe for the video file
		if (!g_pipein) {
			if (!OpenVideo(g_videopath.c_str())) {
				return;
			}
		}
		else {
			// Read a frame from the FFmpeg input pipe
			if (g_pipein && g_pixelBuffer && g_SenderWidth > 0 && g_SenderHeight > 0) {
				if (fread(g_pixelBuffer, 1, g_SenderWidth*g_SenderHeight*4, g_pipein) == 0) {
					// fread = 0 means the end of the file
					// Use the same file and pixel buffer
					fflush(g_pipein);
					_pclose(g_pipein);
					g_pipein = _popen(g_input.c_str(), "rb");
				}
			}
		}
	} // endif video or receiver

	//
	// Draw the received image (approx 2 msec)
	//
	if (g_pixelBuffer) {

		// Must use GetDCEx
		HDC hdc = GetDCEx(g_WorkerHwnd, 0, DCX_WINDOW);
		RECT dr{};
		GetWindowRect(g_WorkerHwnd, &dr);

		BITMAPINFO bmi{};
		ZeroMemory(&bmi, sizeof(BITMAPINFO));
		bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bmi.bmiHeader.biSizeImage = (LONG)(g_SenderWidth * g_SenderHeight * 4); // Pixel buffer size
		bmi.bmiHeader.biWidth = (LONG)g_SenderWidth; // Width of buffer
		bmi.bmiHeader.biHeight = -(LONG)g_SenderHeight; // Height of buffer allowing for bottom up
		bmi.bmiHeader.biPlanes = 1;
		bmi.bmiHeader.biBitCount = 32;
		bmi.bmiHeader.biCompression = BI_RGB;
		// StretchDIBits adapts the pixel buffer to the window size.
		// The sender can be resized or changed.
		// Very fast (< 1msec at 1280x720)
		SetStretchBltMode(hdc, COLORONCOLOR); // Fastest method
		StretchDIBits(hdc,
			0, 0, (dr.right - dr.left), (dr.bottom - dr.top), // destination rectangle
			0, 0, g_SenderWidth, g_SenderHeight, // source rectangle
			g_pixelBuffer,
			&bmi, DIB_RGB_COLORS, SRCCOPY);
		ReleaseDC(g_WorkerHwnd, hdc);

		// Hold at 30 fps reduces CPU load
		receiver.HoldFps(30);

		return;
	}

}


// Initialize the window and tray icon
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	HDC hdc = NULL;
	HWND hwnd = NULL;
	HGLRC hRc = NULL;
	char windowtitle[512]{};

	// save instance handle and create dialog window
	hInst = hInstance;
	HWND hWnd = CreateDialog(hInstance, MAKEINTRESOURCE(IDD_DLG_DIALOG), NULL, (DLGPROC)DlgProc );
	if (!hWnd) {
		return FALSE;
	}
	hWndMain = hWnd;

	// Fill the NOTIFYICONDATA structure and call Shell_NotifyIcon
	ZeroMemory(&niData, sizeof(NOTIFYICONDATA));

	// Get Shell32 version number and set the size of the structure
	ULONGLONG ullVersion = GetDllVersion(_T("Shell32.dll"));

	if(ullVersion >= MAKEDLLVERULL(5, 0,0,0))
		niData.cbSize = sizeof(NOTIFYICONDATA);
	else 
		niData.cbSize = NOTIFYICONDATA_V2_SIZE;

	// Set the ID number (can be anything)
	niData.uID = TRAYICONID;

	// State which structure members are valid
	niData.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP; // Must receive messages
	niData.dwState = 0x00000000; // not hidden
	niData.dwStateMask = NIS_HIDDEN; // alow only the hidden state to be modified - we have no notifications

	// Load the icon
	niData.hIcon = (HICON)LoadImage(hInstance,
									MAKEINTRESOURCE(IDI_STEALTHDLG),
									IMAGE_ICON, 
									GetSystemMetrics(SM_CXSMICON),
									GetSystemMetrics(SM_CYSMICON),
									LR_DEFAULTCOLOR);

	// The window to send messages to and the message to send.
	// Note: the message value should be in the range of WM_APP through 0xBFFF.
	niData.hWnd = hWnd;
    niData.uCallbackMessage = SWM_TRAYMSG;

	wchar_t  ws[512]{};
	wsprintf(ws, L"SpoutWallPaper");
	lstrcpyn(niData.szTip, ws, sizeof(niData.szTip)/sizeof(TCHAR));
	Shell_NotifyIcon(NIM_ADD, &niData);

	// Free the icon handle
	if(niData.hIcon && DestroyIcon(niData.hIcon))
		niData.hIcon = NULL;

	return TRUE;
}

BOOL OnInitDialog(HWND hWnd)
{
	HMENU hMenu = GetSystemMenu(hWnd, FALSE);

	HICON hIcon = (HICON)LoadImage(	hInst,
									MAKEINTRESOURCE(IDI_STEALTHDLG),
									IMAGE_ICON, 
									0, 0, 
									LR_SHARED|LR_DEFAULTSIZE);

	SendMessage(hWnd,WM_SETICON,ICON_BIG,(LPARAM)hIcon);
	SendMessage(hWnd,WM_SETICON,ICON_SMALL,(LPARAM)hIcon);

	return TRUE;
}


// Main right-click context menu for the tray app
void ShowContextMenu(HWND hWnd)
{
	int item = 0;
	POINT pt{};
	char activename[512]{};
	char itemstring[512]{};

	GetCursorPos(&pt);
	HMENU hMenu = CreatePopupMenu();

	if(hMenu) {

		// Insert all the Sender names as menu items
		if(spout.sendernames.GetSenderNames(&Senders)) {
			Senders.clear();
			spout.sendernames.GetSenderNames(&Senders);
            // Add all the Sender names as items to the dialog list.
			if(Senders.size() > 0) {
				// Get the active Sender name
				spout.GetActiveSender(activename);
				item = 0;
				for(iter = Senders.begin(); iter != Senders.end(); iter++) {
					namestring = *iter; // the string to copy
					strcpy_s(name, namestring.c_str());
					// Find it's width and height
					spout.sendernames.getSharedInfo(name, &info);
					sprintf_s(itemstring, "%s : (%d x %d)", name, info.width, info.height);
					InsertMenuA(hMenu, -1, MF_BYPOSITION, WM_APP+item, itemstring);
					// Was it the active Sender ?
					if(strcmp(name, activename) == 0)
						CheckMenuItem (hMenu, WM_APP+item, MF_BYCOMMAND | MF_CHECKED);
					item++;
				} // end menu item loop
				AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
			} // endif Senders size > 0
		} // end go the Sender names OK

		AppendMenu(hMenu, MF_STRING, IDM_VIDEO, _T("Video"));
		AppendMenu(hMenu, MF_STRING, IDM_IMAGE, _T("Image"));
		AppendMenu(hMenu, MF_STRING, IDM_DAILY, _T("Daily"));
		AppendMenu(hMenu, MF_STRING, IDM_SLIDESHOW, _T("Slideshow"));
		AppendMenu(hMenu, MF_STRING, IDM_ABOUT, _T("About"));
		InsertMenu(hMenu, -1, MF_BYPOSITION, SWM_EXIT, _T("Exit"));

		// Note: must set window to the foreground
		// or the menu won't disappear when it should
		SetForegroundWindow(hWnd);

		TrackPopupMenu(	hMenu, 
						TPM_BOTTOMALIGN,
						pt.x, pt.y, 
						0, hWnd, NULL );
		
		DestroyMenu(hMenu);

	}
}

// Get dll version number
ULONGLONG GetDllVersion(LPCTSTR lpszDllName)
{
    ULONGLONG ullVersion = 0;
	HINSTANCE hinstDll = NULL;
    hinstDll = LoadLibrary(lpszDllName);

    if(hinstDll) {
        DLLGETVERSIONPROC pDllGetVersion;
        pDllGetVersion = (DLLGETVERSIONPROC)GetProcAddress(hinstDll, "DllGetVersion");

        if(pDllGetVersion) {
            DLLVERSIONINFO dvi;
            HRESULT hr;
            ZeroMemory(&dvi, sizeof(dvi));
            dvi.cbSize = sizeof(dvi);
            hr = (*pDllGetVersion)(&dvi);
            if(SUCCEEDED(hr))
				ullVersion = MAKEDLLVERULL(dvi.dwMajorVersion, dvi.dwMinorVersion,0,0);
        }
        FreeLibrary(hinstDll);
    }
    return ullVersion;
}

void RestoreWallPaper()
{
	// Get the current wallpaper if not already
	if (g_wallpaperpath.empty()) {
		char szWallPaper[MAX_PATH]{};
		SystemParametersInfoA(SPI_GETDESKWALLPAPER, MAX_PATH, (PVOID)szWallPaper, 0);
		if (szWallPaper) g_wallpaperpath = szWallPaper;
	}

	// Restore original wallpaper if not being shown
	if (!g_wallpaperpath.empty() && !bCurrentWallpaper) {
		SystemParametersInfoA(SPI_SETDESKWALLPAPER, 0, (PVOID)g_wallpaperpath.c_str(), SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
		bCurrentWallpaper = true;
	}
}


bool OpenVideo(std::string filePath)
{
	if (filePath.empty() || _access(filePath.c_str(), 0) == -1)
		return false;

	// Get information from the video file using ffprobe
	// to set the width, height globals, g_SenderWidth and g_SenderHeight
	if (!ffprobe(filePath)) {
		MessageBoxA(NULL, "FFprobe error", "Warning", MB_OK);
		return false;
	}

	if (g_pipein) {
		fflush(g_pipein);
		_pclose(g_pipein);
		g_pipein = nullptr;
	}

	// _popen will open a console window.
	// To hide the output, open a console first and then hide it.
	// An application can have only one console window.
	FILE* pCout = nullptr;
	if (AllocConsole()) freopen_s(&pCout, "CONOUT$", "w", stdout);
	ShowWindow(GetConsoleWindow(), SW_HIDE);

	// Open an input pipe from ffmpeg
	g_input = g_ffmpegPath;
	// Read input at native frame rate 
	g_input += " -re ";
	g_input += " -i ";
	g_input += "\"";
	g_input += filePath;
	g_input += "\"";
	// Specify BGRA pixel format to allow high speed bitmap drawing
	g_input += " -f image2pipe -vcodec rawvideo -pix_fmt bgra -";
	g_pipein = _popen(g_input.c_str(), "rb");
	if (g_pipein) {
		if (g_pixelBuffer) delete[] g_pixelBuffer;
		unsigned int buffersize = g_SenderWidth*g_SenderHeight * 4;
		g_pixelBuffer = new unsigned char[buffersize];
		return true;
	}
	else {
		MessageBoxA(NULL, "FFmpeg open failed", "Warning", MB_OK | MB_TOPMOST);
	}

	return false;
}

void CloseVideo()
{
	if (g_pipein) {
		fflush(g_pipein);
		_pclose(g_pipein);
		g_pipein = nullptr;
	}
	if (g_pixelBuffer) delete[] g_pixelBuffer;
	g_pixelBuffer = nullptr;
	g_SenderWidth = 0;
	g_SenderHeight = 0;
	g_videopath.clear();
}


// Dialog to open video or image file
bool OpenFile(char* filepath, int maxchars, bool bVideo)
{
	OPENFILENAMEA ofn;
	char szFile[MAX_PATH] = "";
	char szDir[MAX_PATH] = "";

	ZeroMemory(&ofn, sizeof(ofn));
	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = NULL; // TODO hwnd;

	// Show the user the existing file in the FileOpen dialog
	if (filepath[0]) {
		strcpy_s(szFile, filepath);
		PathStripPathA(szFile);
	}

	// Set defaults
	if(bVideo)
		ofn.lpstrFilter = "mp4(*.mp4)\0 *.mp4\0mkv(*.mkv)\0 *.mkv\0avi(*.avi)\0 *.avi\0wmv(*.wmv)\0 *.wmv\0All Files (*.*)\0*.*\0";
	else
		ofn.lpstrFilter = "jpg(*.jpg)\0 *.jpg\0png(*.png)\0 *.png\0bmp(*.bmp)\0 *.bmp\0All Files (*.*)\0*.*\0";
	ofn.lpstrDefExt = "";
	ofn.lpstrFile = szFile;
	ofn.nMaxFile = MAX_PATH;
	ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
	if (GetOpenFileNameA(&ofn)) {
		strcpy_s(filepath, maxchars, szFile);
		return true;
	}

	return false;

}


// Dialog to select a slideshow folder
bool OpenFolder(char* filepath, int maxchars)
{
	char szDir[MAX_PATH]{};
	BROWSEINFOA bInfo{};
	bInfo.hwndOwner = NULL; // Owner window
	bInfo.pidlRoot = NULL;
	bInfo.pszDisplayName = szDir; // Address of a buffer to receive the display name of the folder selected by the user
	bInfo.lpszTitle = "Select a folder for slideshow images"; // Title of the dialog
	bInfo.ulFlags = 0;
	bInfo.lpfn = BrowseCallbackProc;
	bInfo.lParam = 0;
	bInfo.iImage = -1;
	// Larger dialog but does not scroll to selected folder
	// bInfo.ulFlags = BIF_NEWDIALOGSTYLE | BIF_NONEWFOLDERBUTTON; // Larger dialog

	// Set starting folder for Browse
	strcpy_s(startingfolder, MAX_PATH, filepath);

	LPITEMIDLIST lpItem = SHBrowseForFolderA(&bInfo);
	if (lpItem != NULL) {
		SHGetPathFromIDListA(lpItem, szDir);
		strcpy_s(filepath, maxchars, szDir);
		return true;
	}

	return false;
}


// Run FFprobe on a video file and produce an ini file with the stream information
bool ffprobe(std::string videoPath)
{
	// Get information from the movie file using ffprobe and write to an ini file
	// Use a batch file with the required ffprobe options and pass the path to ShellExecute
	std::string probepath = "\"";
	probepath += g_exePath;
	probepath += "\\DATA\\FFMPEG\\probe.bat";
	probepath += "\"";

	// Input to ffprobe
	std::string input = "\"";
	input += videoPath;
	input += "\"";

	// In the batch file, %~dp0 returns the Drive and Path to the batch script

	// Open ffprobe and wait for completion
	STARTUPINFOA si = { sizeof(STARTUPINFOA) };
	si = { sizeof(STARTUPINFOA) };
	DWORD dwExitCode = 0;
	ZeroMemory((void*)&si, sizeof(STARTUPINFO));
	si.cb = sizeof(STARTUPINFO);
	si.dwFlags = STARTF_USESHOWWINDOW;
	si.wShowWindow = SW_HIDE; // hide the ffprobe console window
	PROCESS_INFORMATION pi={};
	std::string cmdstring = probepath + " " + input;
	if (CreateProcessA(NULL, (LPSTR)cmdstring.c_str(), NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
		if (pi.hProcess) {
			do {
				GetExitCodeProcess(pi.hProcess, &dwExitCode);
			} while (dwExitCode == STILL_ACTIVE);
			CloseHandle(pi.hProcess);
		}
		if (pi.hThread)	CloseHandle(pi.hThread);
	}
	else {
		MessageBoxA(NULL, "FFprobe CreateProcess failed", "Warning", MB_OK | MB_TOPMOST);
		return false;
	}

	// Read the ini file produced by FFprobe to get the video information
	char initfile[MAX_PATH]={};
	strcpy_s(initfile, MAX_PATH, g_exePath.c_str());
	strcat_s(initfile, MAX_PATH, "\\DATA\\FFMPEG\\myprobe.ini");
	if (_access(initfile, 0) == -1) {
		MessageBoxA(NULL, "FFprobe ini file not found", "Warning", MB_OK | MB_TOPMOST);
		return false;
	}

	// For example :
	// [streams.stream.0]
	// width=1280
	// height = 720

	char tmp[MAX_PATH]={};
	DWORD dwResult = 0;
	g_SenderWidth = 0;
	g_SenderHeight = 0;

	// Find the first video stream
	char stream[32]={};
	for (int i=0; i<10; i++) { // arbritrary maximum
		sprintf_s(stream, 32, "streams.stream.%d", i);
		if (GetPrivateProfileStringA((LPCSTR)stream, (LPSTR)"codec_type", (LPSTR)"0", (LPSTR)tmp, 8, initfile) > 0) {
			if (strcmp(tmp, "video") == 0) {
				if (GetPrivateProfileStringA((LPCSTR)stream, (LPSTR)"width", NULL, (LPSTR)tmp, 8, initfile) > 0)
					g_SenderWidth = atoi(tmp);
				if (GetPrivateProfileStringA((LPCSTR)stream, (LPSTR)"height", NULL, (LPSTR)tmp, 8, initfile) > 0)
					g_SenderHeight = atoi(tmp);
				dwResult = GetPrivateProfileStringA((LPCSTR)"streams.stream.0", (LPSTR)"r_frame_rate", (LPSTR)"30/1", (LPSTR)tmp, 11, initfile);
				if (dwResult == 0)
					dwResult = GetPrivateProfileStringA((LPCSTR)"streams.stream.0", (LPSTR)"avg_frame_rate", (LPSTR)"30/1", (LPSTR)tmp, 11, initfile);
				break;
			}
		}
	}

	if (dwResult > 0) {
		std::string iniValue = tmp;
		auto pos = iniValue.find("/");
		double num = atof(iniValue.substr(0, pos).c_str());
		double den = atof(iniValue.substr(pos + 1, iniValue.npos).c_str());
		if (num > 0 && den > 0) {
			g_FrameRate = (float)(num / den);
		}
	}

	if (g_SenderWidth == 0 || g_SenderHeight == 0)
		return false;

	return true;
}

// Get image files for slideshow
int GetImageFiles(const char* spath, std::vector<std::string>&filenames)
{
	char tmp[MAX_PATH]{};
	HANDLE filehandle = NULL;
	WIN32_FIND_DATAA filedata;

	sprintf_s(tmp, MAX_PATH, "%s\\*.*", spath);
	int nFiles = 0;
	filehandle = FindFirstFileA((LPCSTR)tmp, (LPWIN32_FIND_DATAA)&filedata);
	if (PtrToUint(filehandle) > 0) {
		do {
			if ((filedata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
				if (strstr((char*)filedata.cFileName, ".jpg") != 0
					|| strstr((char*)filedata.cFileName, ".png") != 0
					|| strstr((char*)filedata.cFileName, ".gif") != 0
					|| strstr((char*)filedata.cFileName, ".bmp") != 0
					|| strstr((char*)filedata.cFileName, ".tga") != 0
					|| strstr((char*)filedata.cFileName, ".tif") != 0) {
						slidenames.push_back(filedata.cFileName);
						nFiles++;
				}
			}
		} while (FindNextFileA(filehandle, (LPWIN32_FIND_DATAA)&filedata));
	}

	return nFiles;

} // end GetImageFiles


// Message handler for the app
INT_PTR CALLBACK DlgProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent, ID;
	HMENU hMenu = GetMenu(hWnd);
	char filepath[MAX_PATH]{};
	std::string str;

	switch (message) {

	case SWM_TRAYMSG:

		switch(lParam) {
			case WM_LBUTTONDBLCLK:
				ShowWindow(hWnd, SW_RESTORE);
				break;
			case WM_RBUTTONDOWN:
			case WM_CONTEXTMENU:
				ShowContextMenu(hWnd);
		}
		break;

	case WM_SYSCOMMAND:
		if((wParam & 0xFFF0) == SC_MINIMIZE) {
			ShowWindow(hWnd, SW_HIDE);
			return 1;
		}
		break;

	case WM_COMMAND:
		
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam); 
		ID		= LOWORD(wParam) - WM_APP;

		// ID has the item number of the selected Sender
		if (Senders.size() > 0 && ID >= 0 && ID <= (int)Senders.size()) {
			iter = std::next(Senders.begin(), ID);
			namestring = *iter; // the selected Sender in the list
			strcpy_s(name, namestring.c_str());
			spout.SetActiveSender(name); // make it active
			receiver.SetReceiverName(name); // set the name for the receiver to use
			// Close video
			CloseVideo();
			// Clear slideshow
			slidenames.clear();
			// Disable daily wallpaper display
			bShowDaily = false;
			// Set timer for 30 msec
			KillTimer(hWndMain, 1);
			SetTimer(hWndMain, 1, 30, NULL);
			break;
		}

		switch (wmId) {

			case SWM_EXIT:
			{
				ShowWindow(hWnd, SW_HIDE);
				DestroyWindow(hWnd);
			}
			break;

			case IDM_VIDEO:
				{
					if (OpenFile(filepath, MAX_PATH, true)) {
						// Close existing video
						CloseVideo();
						// Set the new video path
						g_videopath = filepath;
						// Close receiver
						receiver.ReleaseReceiver();
						// Clear any slideshow
						slidenames.clear();
						// Disable daily wallpaper display
						bShowDaily = false;
						// Not showing original wallpaper
						bCurrentWallpaper = false;
						// Set timer for 30 msec
						KillTimer(hWndMain, 1);
						SetTimer(hWndMain, 1, 30, NULL);
					}
					else {
						CloseVideo();
					}
				}
				break;

			case IDM_IMAGE:
			{
				//
				// Download an image for the wallpaper
				//

				// Stop video
				CloseVideo();
				// Close receiver
				receiver.ReleaseReceiver();
				// Clear any slideshow
				slidenames.clear();
				// Default is image not downloaded
				bDailyWallpaper = false;
				if (OpenFile(filepath, MAX_PATH)) {
					// Set the new wallpaper
					SystemParametersInfoA(SPI_SETDESKWALLPAPER, 0, (void*)filepath, SPIF_SENDCHANGE);
					// Save the image path
					g_dailywallpaperpath = filepath;
					// Flag download of wallpaper for exit
					bDailyWallpaper = true;
					// Not showing original wallpaper
					bCurrentWallpaper = false;
					// Bypass Spout and Video in Render()
					bShowDaily = true;
					// Replace Bing daily image details with the image name
					PathStripPathA(filepath);
					copyright = filepath;
					// Set timer for every 2 seconds
					KillTimer(hWndMain, 1);
					SetTimer(hWndMain, 1, 2000, NULL);
				}
			}
			break;

			case IDM_DAILY:
				{
					//
					// Download the Bing daily wallpaper
					//

					// Stop video
					CloseVideo();
					// Close receiver
					receiver.ReleaseReceiver();
					// Clear any slideshow
					slidenames.clear();
					// Default is image not downloaded
					bDailyWallpaper = false;

					// Download the json file describing the daily wallpaper
					std::string jsonFile = g_exePath;
					jsonFile += "\\";
					jsonFile += "\\BingDaily.json";
					const char* url = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US";
					if (URLDownloadToFileA(NULL, url, jsonFile.c_str(), 0, NULL) == S_OK) {
						// Get the json file text as a single string
						if (_access(jsonFile.c_str(), 0) != -1) {
							// Open the file
							std::ifstream logstream(jsonFile);
							// Source file loaded OK ?
							if (logstream.is_open()) {
								std::string logstr = "";
								logstr.assign((std::istreambuf_iterator< char >(logstream)), std::istreambuf_iterator< char >());
								logstr += ""; // ensure a NULL terminator
								logstream.close();
								// Find the image url in the json string
								std::string imageurl = logstr.substr(logstr.find("url")+6);
								imageurl = imageurl.substr(0, imageurl.find("jpg")+3);
								std::string bingUrl = "https://www.bing.com";
								bingUrl += imageurl;
								// Find the title e.g "title":"An ice day to play"
								std::string title = logstr.substr(logstr.find("title")+8);
								title = title.substr(0, title.find("quiz")-3);
								// Remove all '?'
								title.erase(std::remove(title.begin(), title.end(), '?'), title.end());
								// Find the copyright description for About and Exit
								copyright = logstr.substr(logstr.find("copyright")+12);
								copyright = copyright.substr(0, copyright.find("copyrightlink")-3);
								// Make the file name to save
								std::string filePath;
								std::string fileName = g_exePath;
								fileName += "\\DATA\\Images\\";
								fileName += title;
								fileName += ".jpg";
								// Does the image already exist ?
								if (_access(fileName.c_str(), 0) == -1) {
									// If not, download it
									if (URLDownloadToFileA(NULL, bingUrl.c_str(), fileName.c_str(), 0, NULL) == S_OK) {
										// Set the image path
										filePath = fileName;
									}
								}
								else {
									filePath = fileName;
								}

								if (!filePath.empty()) {
									// Set the new wallpaper
									SystemParametersInfoA(SPI_SETDESKWALLPAPER, 0, (void*)filePath.c_str(), SPIF_SENDCHANGE);
									// Save the daily wallpaper image path
									g_dailywallpaperpath = filePath;
									// Not showing original wallpaper
									bCurrentWallpaper = false;
									// Flag download of wallpaper for exit
									bDailyWallpaper = true;
									// Bypass Spout and Video in Render()
									bShowDaily = true;
									// Set timer for every 2 seconds
									KillTimer(hWndMain, 1);
									SetTimer(hWndMain, 1, 2000, NULL);

								}
							}
							remove(jsonFile.c_str());
						}
					}
				}
				break;

			case IDM_SLIDESHOW:
			{
				char path[MAX_PATH]{};
				char name[MAX_PATH]{};
				strcpy_s(path, MAX_PATH, g_slideshowpath); // starting folder
				if (OpenFolder(path, MAX_PATH)) {
					slidenames.clear();
					if (GetImageFiles(path, slidenames) > 0) {
						// Allow selection of slide duration
						if (SelectSlideDuration()) {
							// Close video
							CloseVideo();
							// Close receiver
							receiver.ReleaseReceiver();
							// Disable daily wallpaper display
							bShowDaily = false;
							// Save selected folder
							strcpy_s(g_slideshowpath, MAX_PATH, path);
							// Write to the registry for the next start
							WritePathToRegistry(HKEY_CURRENT_USER, "Software\\Leading Edge\\SpoutWallpaper", "slideshowfolder", g_slideshowpath);
							// TODO - user select slide time
							WriteDwordToRegistry(HKEY_CURRENT_USER, "Software\\Leading Edge\\SpoutWallpaper", "slideshowtime", g_slideshowtime);
							// Reset counter and timer
							nCurrentImage = 0;
							// Set start time
							g_start = ElapsedMicroseconds()/1000.0;
							// Set timer for every 1 second
							KillTimer(hWndMain, 1);
							SetTimer(hWndMain, 1, 1000, NULL);
						}
						else {
							slidenames.clear();
							nCurrentImage = 0;
							g_start = 0.0;
						}
					}
					else {
						SpoutMessageBox(NULL, "No image files in the folder", "SpoutWallPaper", MB_ICONWARNING | MB_OK);
					}
				}
				else {
					slidenames.clear();
					nCurrentImage = 0;
					g_start = 0.0;
				}
			}
			break;
		
			case IDM_ABOUT:
			{
				HICON hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_STEALTHDLG));
				SpoutMessageBoxIcon(hIcon);
				std::string str = "Version 1.003\n\nChange Windows wallpaper\nSpout sender, video, image, Bing daily, slideshow.\n\n";
				str +=	"<a href=\"https://spout.zeal.co/\">https://spout.zeal.co/</a>\n";
				if (!copyright.empty()) {
					str += "\nWallpaper image\n";
					str += copyright;
					str += "\n\n";
				}
				SpoutMessageBox(NULL, str.c_str(), " ", MB_USERICON | MB_OK, "SpoutWallPaper");
			}
			break;

		}

		return 1;

	case WM_INITDIALOG:
		return OnInitDialog(hWnd);

	case WM_CLOSE:
		DestroyWindow(hWnd);
		break;

	case WM_DESTROY:
		niData.uFlags = 0;
		Shell_NotifyIcon(NIM_DELETE, &niData);
		PostQuitMessage(0);
		break;
	}
	return 0;
}


// Select slide duration aft selecting folder
bool SelectSlideDuration()
{
	// 2, 4, 10, 30 seconds
	// 1, 2,  4, 10 minutes
	// 1, 2,  5, 10 hours
	switch (g_slideshowtime) { // seconds
		case     2: durationindex =  0; break; //  2 seconds
		case     4: durationindex =  1; break; //  4 seconds
		case    10: durationindex =  2; break; // 10 seconds
		case    30: durationindex =  3; break; // 30 seconds
		case    60: durationindex =  4; break; //  1 minute
		case   120: durationindex =  5; break; //  2 minutes
		case   240: durationindex =  6; break; //  4 minutes
		case   600: durationindex =  7; break; // 10 minutes
		case  1800: durationindex =  8; break; // 30 minutes
		case  3600: durationindex =  9; break; //  1 hour (60 minutes)
		case  7200: durationindex = 10; break; //  2 hours
		case 14400: durationindex = 11; break; //  4 hours
		case 36000: durationindex = 12; break; // 10 hours
		default: break;
	}
	// Show the dialog box 
	if((int)DialogBoxA(hInst, MAKEINTRESOURCEA(IDD_DURATIONBOX), g_hWnd, (DLGPROC)SlideDuration) != 0) {
		// OK - new duration has been selected
		g_slideshowtime = (DWORD)duration[durationindex]; // seconds
		return true;
	}

	return false;

}

// Message handler for slide duration selection
INT_PTR CALLBACK SlideDuration(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	HWND hwndList = NULL;
	char formats[13][128] = { "2 seconds", "4 seconds", "10 seconds", "30 seconds", "1 minute", "2 minutes", "4 minutes", "10 minutes", "30 minutes", "1 hour", "2 hours", "4 hours", "10 hours" };

	switch (message) {

	case WM_INITDIALOG:
		SetWindowTextA(hDlg, "Slide duration");
		hwndList = GetDlgItem(hDlg, IDC_SLIDE_DURATION);
		SendMessageA(hwndList, (UINT)CB_RESETCONTENT, 0, 0L);
		for (int k = 0; k < 13; k++)
			SendMessageA(hwndList, (UINT)CB_ADDSTRING, 0, (LPARAM)formats[k]);
		SendMessageA(hwndList, CB_SETCURSEL, (WPARAM)durationindex, (LPARAM)0);

		// Random
		if (bRandom)
			CheckDlgButton(hDlg, IDC_SLIDE_RANDOM, BST_CHECKED);
		else
			CheckDlgButton(hDlg, IDC_SLIDE_RANDOM, BST_UNCHECKED);

		return (INT_PTR)TRUE;

	case WM_COMMAND:
		// Combo box selection
		if (HIWORD(wParam) == CBN_SELCHANGE) {
			if ((HWND)lParam == GetDlgItem(hDlg, IDC_SLIDE_DURATION)) {
				durationindex = static_cast<unsigned int>(SendMessageA((HWND)lParam, (UINT)CB_GETCURSEL, (WPARAM)0, (LPARAM)0));
			}
		}


		// Drop though if not selected
		switch (LOWORD(wParam)) {
		case IDOK:
			// Random slides 
			if (IsDlgButtonChecked(hDlg, IDC_SLIDE_RANDOM) == BST_CHECKED)
				bRandom = true;
			else
				bRandom = false;
			EndDialog(hDlg, 1);
			return (INT_PTR)TRUE;
		case IDCANCEL:
			EndDialog(hDlg, 0);
			return (INT_PTR)TRUE;
		default:
			return (INT_PTR)FALSE;
		}
		break;
	}

	return (INT_PTR)FALSE;
}


// Used by the OpenFolder function to choose a folder
INT CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM lp, LPARAM pData)
{
	TCHAR szDir[MAX_PATH];
	TCHAR initialDir[MAX_PATH];

	mbstowcs_s(NULL, &initialDir[0], MAX_PATH, &startingfolder[0], MAX_PATH);

	switch (uMsg)
	{
	case BFFM_INITIALIZED:
		_tcscpy_s(szDir, MAX_PATH, initialDir);
		// Not working for modern interface
		SendMessage(hwnd, BFFM_SETEXPANDED, TRUE, (LPARAM)szDir);
		SendMessage(hwnd, BFFM_SETSELECTION, TRUE, (LPARAM)szDir);
		break;
	}

	return 0;
}

// To find the worker window
BOOL CALLBACK EnumWindowsProc(HWND hWnd, LPARAM lParam)
{
	HWND hwnd = FindWindowEx(hWnd, NULL, L"SHELLDLL_DefView", NULL);
	if (hwnd) {
		HWND* ret = (HWND*)lParam;
		*ret = FindWindowEx(NULL, hWnd, L"WorkerW", NULL);
	}
	return TRUE;
}

// there is no more

