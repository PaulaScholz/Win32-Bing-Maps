// GraphicsTestWin32.cpp : Downloads and displays a Bing Map.  April 2020
//
#include <atlbase.h>
#include "framework.h"
#include "GraphicsTestWin32.h"
#include <wininet.h>
#include <initguid.h>
#include <atlstr.h>
#include <vector>
#include <wincodec.h>
#include <wincodecsdk.h>
#pragma comment(lib, "WindowsCodecs.lib")
#pragma comment(lib, "wininet.lib")

#define MAX_LOADSTRING 100
#define MAX_DEBUGMSG 256

// define some common COM macros that make life simpler
#pragma warning(disable : 4127)  // conditional expression is constant

// Macro that calls a COM method returning HRESULT value.
#define CHK_HR(stmt)        do { hr=(stmt); if (FAILED(hr)) goto CleanUp; } while(0)

// Macro to verify memory allcation.
#define CHK_ALLOC(p)        do { if (!(p)) { hr = E_OUTOFMEMORY; goto CleanUp; } } while(0)

// Macro that releases a COM object if not NULL.
#define SAFE_RELEASE(p)     do { if ((p)) { (p)->Release(); (p) = NULL; } } while(0)

// used for calculating scanline stride
#define DIB_WIDTHBYTES(bits) ((((bits) + 31)>>5)<<2)

// current UI state
enum class CurrentUIState
{
    START,			// Start, display instructions
    SEATTLE,		// Display Seattle map
    PORTLAND,		// Display Portland Map
    SANFRAN,		// Display San Francisco map
};

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name
WCHAR szDebugMsg[MAX_DEBUGMSG];                 // for OutputDebugMsg

CurrentUIState		g_uiState = CurrentUIState::START;

// default size of UDP packets and DNS payloads, so it
// makes for a convenient InternetReadFile buffer size
const int			g_nPacketSize = 512;

// this structure is used to hold a pointer 
// to a memory buffer and its current size count.
// These structures are stored in a vector as a 
// file is read from the Internet with InternetReadFile
// in discrete chunks.  iovcnt will always be 512 bytes
// (or the size of the g_nPacketSize constant), except
// for the last IOVEC in the vector, which will contain
// the number of bytes from the last Read operation.
//
// Used in GetBingMap to hold pointers to chunks of a
// .jpg file as it is downloaded.
typedef struct iovec
{
    LPBYTE	pBytes;
    int		iovcnt;

} IOVEC, * PIOVEC;

// three HBITMAPs for three cities, created in 
// GetMap, painted in DisplayMap, selected
// in the WM_COMMAND message handler, and destroyed
// in DestroyGDIObjects() called from WM_DESTROY handler.
HBITMAP				g_hbmSeattle;
HBITMAP				g_hbmPortland;
HBITMAP				g_hbmSanFran;

// some fonts for writing to the screen, created in
// CreateSmallUserSizedFonts(), painted by DrawText in
// DisplayInstructions, and destroyed in DestroyGDIObjects().
HFONT g_hFontSmallBold,
g_hFontSmallNormal,
g_hOldFont;

// Created in InitInstance, used in GetMap
// to create Image objects from a memory buffer.
// It is global because it is a COM server "singleton"
// commonly used in more than one function.  It is
// released in the WM_DESTROY message handler.
IWICImagingFactory* g_pIWICFactory = NULL;

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);
void DestroyGDIObjects();
HRESULT GetBingMap(LPCTSTR pszCityName, HBITMAP& refHbmOut, int requestedWidth, int requestedHeight);
void CreateSmallUserSizedFonts();
LRESULT DisplayInstructions(HWND hWnd);
LRESULT DisplayMap(HWND hWnd, HBITMAP& hbmCity);

// Entry point
int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_GRAPHICSTESTWIN32, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_GRAPHICSTESTWIN32));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int) msg.wParam;
}

//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_GRAPHICSTESTWIN32));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_GRAPHICSTESTWIN32);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   /*************  Make sure you do this for the Windows Imaging Component   ******************/

    // initialize COM for Windows Imaging Component.  Make sure you use COINIT_MULTITHREADED.
   CoInitializeEx(NULL, COINIT_MULTITHREADED);

   // create the WICImagingFactory
   HRESULT hr = CoCreateInstance(
       CLSID_WICImagingFactory,
       NULL,
       CLSCTX_INPROC_SERVER,
       IID_PPV_ARGS(&g_pIWICFactory));


   /*************  Make sure you do this for Windows Imaging Component   ******************/

   if (S_OK != hr)
   {
       MessageBox(NULL,
           (LPCTSTR)TEXT("Could not create WICImagingFactory!"),
           (LPCTSTR)TEXT("Err! - InitInstance()"),
           MB_OK | MB_ICONEXCLAMATION);

       return FALSE;
   }
   
   // does just what it says
   CreateSmallUserSizedFonts();

   g_hbmSeattle = NULL;
   g_hbmPortland = NULL;
   g_hbmSanFran = NULL;

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE: Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Parse the menu selections:
            switch (wmId)
            {
            case ID_CITY_SEATTLE:

                // only get the map from the Internet once
                if (NULL == g_hbmSeattle)
                {
                    GetBingMap(L"Seattle", g_hbmSeattle, 800, 500);
                }

                // set our paint state to Seattle
                g_uiState = CurrentUIState::SEATTLE;

                // trigger a repaint
                InvalidateRect(hWnd, NULL, TRUE);
                UpdateWindow(hWnd);

                break;

            case ID_CITY_PORTLAND:

                // only get the map from the Internet once
                if (NULL == g_hbmPortland)
                {
                    GetBingMap(L"Portland", g_hbmPortland, 600, 600);
                }

                // set our paint state to Portland
                g_uiState = CurrentUIState::PORTLAND;

                // trigger a repaint
                InvalidateRect(hWnd, NULL, TRUE);
                UpdateWindow(hWnd);

                break;

            case ID_CITY_SANFRANCISCO:

                // only get the map from the Internet once
                if (NULL == g_hbmSanFran)
                {
                    GetBingMap(L"San Francisco", g_hbmSanFran, 500, 400);
                }

                // set our paint state to San Francisco
                g_uiState = CurrentUIState::SANFRAN;

                // trigger a repaint
                InvalidateRect(hWnd, NULL, TRUE);
                UpdateWindow(hWnd);

                break;

            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;

            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;

            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;

    case WM_PAINT:
        {
            switch (g_uiState)
            {
            case CurrentUIState::START:
                DisplayInstructions(hWnd);
                break;
            case CurrentUIState::SEATTLE:
                DisplayMap(hWnd, g_hbmSeattle);
                break;
            case CurrentUIState::PORTLAND:
                DisplayMap(hWnd, g_hbmPortland);
                break;
            case CurrentUIState::SANFRAN:
                DisplayMap(hWnd, g_hbmSanFran);
                break;
            }
        }
        break;

    case WM_DESTROY:

        // delete the fonts and city bitmap objects
        DestroyGDIObjects();

        // destroy the global WIC Factory
        SAFE_RELEASE(g_pIWICFactory);

        // shut down COM
        CoUninitialize();

        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }

    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// create the two small user-sized fonts.  This is done in InitInstance.
void CreateSmallUserSizedFonts()
{
    LOGFONT	lfSmallBold, lfSmallNormal;

    // zero out the logfont structures
    ZeroMemory(&lfSmallBold, sizeof(lfSmallBold));
    ZeroMemory(&lfSmallNormal, sizeof(lfSmallNormal));

    // delete previous fonts
    DeleteObject(g_hFontSmallBold);
    DeleteObject(g_hFontSmallNormal);

    // Font size
    int nDesiredPointSize = -36;

    nDesiredPointSize /= 100;		// divide by 100.  not in documentation

    // create the bold font used for titles, Cleartype, Tahoma
    lfSmallBold.lfHeight = nDesiredPointSize;
    lfSmallBold.lfWidth = 0;
    lfSmallBold.lfEscapement = 0;
    lfSmallBold.lfOrientation = 0;
    lfSmallBold.lfWeight = FW_BOLD;
    lfSmallBold.lfItalic = FALSE;
    lfSmallBold.lfUnderline = FALSE;
    lfSmallBold.lfStrikeOut = 0;
    lfSmallBold.lfCharSet = ANSI_CHARSET;
    lfSmallBold.lfOutPrecision = OUT_DEFAULT_PRECIS;
    lfSmallBold.lfClipPrecision = CLIP_DEFAULT_PRECIS;
    lfSmallBold.lfQuality = CLEARTYPE_QUALITY;
    lfSmallBold.lfPitchAndFamily = DEFAULT_PITCH | FF_SWISS;
    wcscpy_s(lfSmallBold.lfFaceName, TEXT("Segoe UI"));

    g_hFontSmallBold = CreateFontIndirect(&lfSmallBold);

    // create the normal font used for titles, Cleartype, Tahoma
    lfSmallNormal.lfHeight = nDesiredPointSize;
    lfSmallNormal.lfWidth = 0;
    lfSmallNormal.lfEscapement = 0;
    lfSmallNormal.lfOrientation = 0;
    lfSmallNormal.lfWeight = FW_NORMAL;
    lfSmallNormal.lfItalic = FALSE;
    lfSmallNormal.lfUnderline = FALSE;
    lfSmallNormal.lfStrikeOut = 0;
    lfSmallNormal.lfCharSet = ANSI_CHARSET;
    lfSmallNormal.lfOutPrecision = OUT_DEFAULT_PRECIS;
    lfSmallNormal.lfClipPrecision = CLIP_DEFAULT_PRECIS;
    lfSmallNormal.lfQuality = CLEARTYPE_QUALITY;
    lfSmallNormal.lfPitchAndFamily = DEFAULT_PITCH | FF_SWISS;
    wcscpy_s(lfSmallNormal.lfFaceName, TEXT("Segoe UI"));

    g_hFontSmallNormal = CreateFontIndirect(&lfSmallNormal);
}

// destroy global GDI objects to avoid a memory leak
void DestroyGDIObjects()
{
    DeleteObject(g_hFontSmallBold);
    DeleteObject(g_hFontSmallNormal);
    DeleteObject(g_hOldFont);			// just in case

    DeleteObject((HGDIOBJ)g_hbmSeattle);
    DeleteObject((HGDIOBJ)g_hbmPortland);
    DeleteObject((HGDIOBJ)g_hbmSanFran);
}

LRESULT DisplayInstructions(HWND hWnd)
{
    RECT rect;
    HRGN hrgnClip;
    HDC hdc;
    PAINTSTRUCT ps;
    TEXTMETRIC tm;

    // this hdc will be destroyed on EndPaint
    hdc = BeginPaint(hWnd, &ps);

    // The rectangle we get back is in Desktop coordinates, so we need to
    //   modify it to reflect coordinates relative to this window.
    //
    GetWindowRect(hWnd, &rect);

    rect.right -= rect.left;
    rect.bottom -= rect.top;
    rect.left = rect.top = 0;

    // create a region based on the extent of the current window
    hrgnClip = CreateRectRgnIndirect(&ps.rcPaint);

    // select that region into our device context
    SelectClipRgn(hdc, hrgnClip);

    // select the small bold font into the device context
    g_hOldFont = (HFONT)SelectObject(hdc, g_hFontSmallBold);

    // get the text metric data about the selected font
    GetTextMetrics(hdc, &tm);

    // compute the pixel length of the string so we can
    // center it on the screen

    SIZE charSize;
    RECT rectText;

    CString aString = L"Select a city from City menu.";

    GetTextExtentPoint32(hdc, (LPCTSTR)aString, aString.GetLength(), &charSize);

    // centered at 50% of the client area width
    int nStartPointX = (rect.right - charSize.cx) / 2;

    // start the title at 50% down from the top of the client area
    int nStartPointY = rect.bottom / 2;

    SetRect(&rectText, nStartPointX, nStartPointY,
        nStartPointX + charSize.cx, nStartPointY + charSize.cy);

    DrawText(hdc, (LPCTSTR)aString, aString.GetLength(), &rectText, DT_BOTTOM | DT_SINGLELINE | DT_CENTER | DT_NOCLIP);

    // don't delete the fonts as they are global and are deleted in DestroyGDIObjects()
    SelectObject(hdc, g_hOldFont);

    // deselect the clip region and delete it
    SelectClipRgn(hdc, NULL);
    DeleteObject(hrgnClip);

    // this frees the hdc created with BeginPaint
    EndPaint(hWnd, &ps);

    return 0;
}

// paint the map on the screen
LRESULT DisplayMap(HWND hWnd, HBITMAP& hbmCity)
{
    RECT rect;
    HRGN hrgnClip;
    HDC hdc;
    HDC	hMemDC;
    BITMAP bm;
    HBITMAP hbmOld;
    PAINTSTRUCT ps;

    // this hdc will be destroyed on EndPaint
    hdc = BeginPaint(hWnd, &ps);

    // The rectangle we get back is in Desktop coordinates, so we need to
    //   modify it to reflect coordinates relative to this window.
    //
    GetWindowRect(hWnd, &rect);

    rect.right -= rect.left;
    rect.bottom -= rect.top;
    rect.left = rect.top = 0;

    // create a region based on the extent of the current window
    hrgnClip = CreateRectRgnIndirect(&ps.rcPaint);

    // select that region into our device context
    SelectClipRgn(hdc, hrgnClip);

    // create a memory device context into which we can select the
    // interface images for blitting onto the window instance
    hMemDC = CreateCompatibleDC(hdc);

    // select the City bitmap into the memory device context so we can blit it
    hbmOld = (HBITMAP)SelectObject(hMemDC, hbmCity);

    // create a brush with a deep sky blue color for the background
    HBRUSH hBrush = CreateSolidBrush(RGB(0, 191, 255)); // background color brush, deep sky blue

    // select the brush into the paint dc
    HBRUSH hBrushOld = (HBRUSH)SelectObject(hdc, (HGDIOBJ)hBrush);

    // fill the paint hdc background with the background color
    FillRect(hdc, &rect, hBrush);

    // remove the brush from the hdc by selecting the "old" one
    SelectObject(hdc, (HGDIOBJ)hBrushOld);

    // done with the background color brush, delete it
    DeleteObject((HGDIOBJ)hBrush);

    // get the size of the hbmCity bitmap
    GetObject(hbmCity, sizeof(bm), &bm);

    // compute coordinates to blit in the center of our client area
    int nxDest = (rect.right - bm.bmWidth) / 2;
    int nyDest = (rect.bottom - bm.bmHeight) / 2;

    // now, blit the composed hMemDC bitmap onto the paint dc
    BitBlt(
        hdc,
        nxDest,
        nyDest,
        bm.bmWidth,
        bm.bmHeight,
        hMemDC,
        0,
        0,
        SRCCOPY);

    // unselect the city bitmap image from the hMemDC
    SelectObject(hMemDC, (HGDIOBJ)hbmOld);

    // deselect the clip region and delete it
    SelectClipRgn(hdc, NULL);
    DeleteObject(hrgnClip);

    // get rid of the hMemDC, we no longer need it
    DeleteDC(hMemDC);

    // this frees the hdc created with BeginPaint
    EndPaint(hWnd, &ps);

    // notice that we do not destroy the hbmCity HBITMAP, but keep
    // it from call to call. It is destroyed in DestroyGDIObjects(),
    // called from the WM_DESTROY message handler.
    return 0;
}

HRESULT GetBingMap(LPCTSTR pszCityName, HBITMAP& refHbmOut, int requestedWidth = 500, int requestedHeight = 400)
{
    HINTERNET hInternet = NULL;
    HINTERNET hMapUrl = NULL;
    DWORD     dwBytesRead = 0;
    HRESULT	  hr = S_OK;

    DWORD dwStatus = 0;
    DWORD dwIndex = 0;
    HANDLE hOpen = NULL;

    // pointer to a WICStream interface
    IWICStream* pIWICStream = NULL;

    // pointer to a WICBitmapDecoder interface
    IWICBitmapDecoder* pIWICDecoder = NULL;

    // pointer to a WICBitmapDecoderFrame interface
    IWICBitmapFrameDecode* pIWICBitmapFrameDecode = NULL;

    // pointer to a WICFormatConverter interface
    IWICFormatConverter* pIWICConvertedFrame = NULL;

    // these are the Bing Maps defaults
    const int defaultMapWidth = 500;
    const int defaultMapHeight = 400;

    UINT retrievedWidth = 0;
    UINT retrievedHeight = 0;

    // Array used to initially hold the chunks of data
    // read from the internet. This is allocated on the stack.
    LPBYTE sBuf[g_nPacketSize];

    // a contiguous byte array used to hold all downloaded data
    // made up of a number of copied sBuf arrays
    LPBYTE pBuf = NULL;

    // default to Seattle, naturally. Best in the west.
    if (NULL == pszCityName)
    {
        pszCityName = L"Seattle";
    }

    if (requestedWidth <= 50)
    {
        requestedWidth = defaultMapWidth;
    }

    if (requestedHeight <= 50)
    {
        requestedHeight = defaultMapHeight;
    }

    // build a URL for the call to Bing Maps
    // get a Bing Maps Key
    // https://docs.microsoft.com/en-us/bingmaps/getting-started/bing-maps-dev-center-help/getting-a-bing-maps-key

    // Insert your Bing Maps key here
    CString strBingMapsKey = TEXT("Your Bing Maps Key Here");

    // this query will return a .jpg image
    // https://docs.microsoft.com/en-us/bingmaps/rest-services/imagery/get-a-static-map
    CString strMapUrl = TEXT("https://dev.virtualearth.net/REST/v1/Imagery/Map/AerialWithLabels/");
    
    CString strWidth;
    CString strHeight;

    strWidth.Format(L"?mapSize=%d,", requestedWidth);
    strHeight.Format(L"%d&key=", requestedHeight);

    strMapUrl.Append(pszCityName);
    strMapUrl.Append(strWidth);
    strMapUrl.Append(strHeight);
    strMapUrl.Append(strBingMapsKey);

    // open the WinInet stuff
    hInternet = InternetOpen(L"GraphicsTestWin32", INTERNET_OPEN_TYPE_DIRECT, NULL, NULL, 0);

    if (hInternet)
    {

        // open the Bing Maps URL
        hMapUrl = InternetOpenUrl(hInternet, (LPCTSTR)strMapUrl, NULL, -1L,
            INTERNET_FLAG_RELOAD |
            INTERNET_FLAG_PRAGMA_NOCACHE |
            INTERNET_FLAG_NO_CACHE_WRITE,
            NULL);

        // if we opened the Bing Maps URL, then read the map data
        if (hMapUrl)
        {
            // success, 
            OutputDebugString(L"Bing Maps HTTP call successful!\n");

            hr = S_OK;

            ::ZeroMemory(&sBuf, g_nPacketSize * sizeof(BYTE));

            // a vector of buffer pointer objects we use to hold
            // pointers to dynamically allocated memory for reading graphics
            // from the Internet or anything else where we don't know the
            // size in advance.
            std::vector<IOVEC>	readBufferArray;

            BOOL bRead = InternetReadFile(hMapUrl, &sBuf, g_nPacketSize, &dwBytesRead);

            // this structure will hold a pointer to a block of memory read from
            // the internet, and the block's size in bytes
            IOVEC FileMemoryBlock{};

            // read the map jpg in chunks
            while (bRead && (dwBytesRead > 0))
            {
                // allocate a new block of memory for the iovec
                LPBYTE pFMBBuf = NULL;

                // this little block of heap memory goes into our iovec
                CHK_ALLOC(pFMBBuf = new BYTE[dwBytesRead]);

                // store the number of bytes read in the IOVEC structure
                FileMemoryBlock.iovcnt = dwBytesRead;

                // copy the bytes read to the iovec byte buffer
                memcpy(pFMBBuf, sBuf, (size_t)dwBytesRead);

                // store the pointer to the byte buffer in the IOVEC structure
                FileMemoryBlock.pBytes = pFMBBuf;

                // put a copy of the IOVEC structure in the vector
                readBufferArray.push_back(FileMemoryBlock);

                // read the file again for the next block
                bRead = InternetReadFile(hMapUrl, &sBuf, g_nPacketSize, &dwBytesRead);

            }  // end while

            // close the handle on the HTTP request
            InternetCloseHandle(hMapUrl);

            if (readBufferArray.size() > 0)
            {
                // the size of the last packet is already in the IOVEC structure,
                // FileMemoryBlock, so use it
                size_t tBufSize = ((readBufferArray.size() - 1) * g_nPacketSize) +
                    FileMemoryBlock.iovcnt;

                // allocate a contiguous buffer the size of our segments
                CHK_ALLOC(pBuf = new BYTE[tBufSize]);

                // iterate through the vector and copy the memory segments
                // into the final buffer
                std::vector<IOVEC>::iterator it;
                size_t i = 0;

                // the readBufferArray contents of IOVEC structures will be 
                // cleared when it goes out of scope.
                for (it = readBufferArray.begin(); it < readBufferArray.end(); it++)
                {
                    // we're not allocating any memory here, we're just incrementing a pointer
                    // to a segment of memory inside our contiguous pBuf buffer
                    LPBYTE pNew = pBuf + i;

                    // copy the memory in the iovec segment to the pBuf buffer.  Remember, pNew points
                    // to an address inside already-allocated pBuf.
                    memcpy(pNew, (const void*)(it->pBytes), (size_t)(it->iovcnt));

                    // increment the buffer pointer inside pBuf
                    i += (size_t)(it->iovcnt);

                    // recover the heap memory from the iovec segment, we no longer need it
                    // because it has been copied into the pBuf
                    delete[] it->pBytes;
                }


                /************************************************************************/
                // this is how we would have done this on WindowsCE, but on the desktop
                // we need to do something quite different.
                //IImage* pImage = NULL;

                //ImageInfo imageInfoObject;

                //// now, the image bits are held in pBuf, so make an image from them and put it in pImage
                //g_pImageFactory->CreateImageFromBuffer((const void*)pBuf, (UINT)i, BufferDisposalFlagNone, &pImage);
                /************************************************************************/

                // first, we need to put the newly read bytes in a WICStream object                
                CHK_HR(g_pIWICFactory->CreateStream(&pIWICStream));

                // documentation says this is dangerous, but in this case we
                // already have the bytes and they'll be valid for the duration
                // of this method, so go ahead, live dangerously!  Not really, it's quite safe in this case.
                // https://docs.microsoft.com/en-us/windows/win32/api/wincodec/nf-wincodec-iwicstream-initializefrommemory
                CHK_HR(pIWICStream->InitializeFromMemory(pBuf, (DWORD)tBufSize));

                // make a Bitmap decoder from the stream
                CHK_HR(g_pIWICFactory->CreateDecoderFromStream(
                    pIWICStream,                    // The stream to use to create the decoder
                    NULL,                           // Do not prefer a particular codec vendor
                    WICDecodeMetadataCacheOnLoad,   // Cache metadata when needed
                    &pIWICDecoder));                // Pointer to the decoder

                // the number of frames in this image. For JPEGs, it should be 1 only
                UINT nCount = 0;

                CHK_HR(pIWICDecoder->GetFrameCount(&nCount));

                if (nCount >= 1)
                {
                    // get the frame. JPEGs have only one
                    CHK_HR(pIWICDecoder->GetFrame(0, &pIWICBitmapFrameDecode));

                    // retrieve the image dimensions in case Bing sent us a
                    // different size from what we requested
                    CHK_HR(pIWICBitmapFrameDecode->GetSize(&retrievedWidth, &retrievedHeight));

                    // to convert the format of the image from JPEG, we need to create a converter
                    CHK_HR(g_pIWICFactory->CreateFormatConverter(&pIWICConvertedFrame));

                    // convert the frame to 32bppBGR
                    CHK_HR(pIWICConvertedFrame->Initialize(
                        pIWICBitmapFrameDecode,         // frame to convert
                        GUID_WICPixelFormat32bppBGR,  // desired pixel format
                        WICBitmapDitherTypeNone,        // no dithering
                        NULL,                           // desired palette
                        0.f,                            // alpha threshold percent
                        WICBitmapPaletteTypeCustom      // palette translation type
                    ));

                    // no need to resize the frame, we requested it from Bing Maps at the size we want

                    // Render the image to a GDI device context
                    HBITMAP hDIBBitmap = NULL;

                    try
                    {
                        // Get a DC for the full screen
                        HDC hdcScreen = GetDC(NULL);
                        if (!hdcScreen)
                        {
                            throw 1;
                        }                            

                        BITMAPINFO bminfo;
                        ZeroMemory(&bminfo, sizeof(bminfo));
                        bminfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
                        bminfo.bmiHeader.biWidth = (LONG)retrievedWidth;
                        bminfo.bmiHeader.biHeight = -(LONG)retrievedHeight;
                        bminfo.bmiHeader.biPlanes = 1;
                        bminfo.bmiHeader.biBitCount = 32;
                        bminfo.bmiHeader.biCompression = BI_RGB;

                        void* pvImageBits = nullptr;    // Freed in DestroyGDIObjects

                        // Create the HBITMAP. Now, the hDIBBitmap points to the image buffer
                        hDIBBitmap = CreateDIBSection(hdcScreen, &bminfo, DIB_RGB_COLORS, &pvImageBits, NULL, 0);

                        if (!hDIBBitmap)
                        {
                            throw 2;
                        }                           

                        ReleaseDC(NULL, hdcScreen);

                        // Calculate the number of bytes in 1 scanline
                        UINT nStride = DIB_WIDTHBYTES(retrievedWidth * 32);

                        // Calculate the total size of the image
                        UINT numberOfImageBytes = nStride * retrievedHeight;

                        // Copy the converted frame pixels to the DIB section image buffer
                        CHK_HR(pIWICConvertedFrame->CopyPixels(nullptr, nStride, numberOfImageBytes, reinterpret_cast<LPBYTE>(pvImageBits)));

                        // The refHBitmapOut must have DeleteObject called on it somewhere or
                        // there will be a GDI memory and handle leak!!!!!!
                        //
                        // We do this in DestroyGDIObjects, called from the WM_DESTROY event handler.
                        refHbmOut = hDIBBitmap;
                            
                    }
                    catch (...)
                    {
                        OutputDebugString(L"Error creating refHbmOut.\n");
                        throw;
                    }
                }
                else
                {
                    OutputDebugString(L"No image frames in decoder.\n");

                    goto CleanUp;
                }               
            } //endif g_readBufferArray.size() > 0
        }  // endif hMapUrl
        else
        {
            // Internet failure, report why 
            LPTSTR lpExtended;
            DWORD dwLength = 0;
            DWORD dwError;

            hr = E_FAIL;

            // find the length of the error
            if (!InternetGetLastResponseInfo(&dwError, NULL, &dwLength) &&
                GetLastError() == ERROR_INSUFFICIENT_BUFFER)
            {
                // allocate a buffer long enough to handle the error, plus 1
                CHK_ALLOC(lpExtended = (LPTSTR)LocalAlloc(LPTR, dwLength + 1));

                // get the error text
                InternetGetLastResponseInfo(&dwError, lpExtended, &dwLength);

                // write it to the debug console
                _snwprintf_s(szDebugMsg, MAX_DEBUGMSG, L"GetMap InternetOpenUrl Error: %s\n", lpExtended);
                OutputDebugString(szDebugMsg);

                // free the error text memory
                LocalFree(lpExtended);
            }
            else
            {
                OutputDebugString(L"Warning: GetMap InternetOpenUrl extended error reported with no response info.\n");
            }

        } // hMapUrl

    } // hInternet
    else
    {
        OutputDebugString(L"Error: Could not get HINTERNET handle\n");
        hr = E_FAIL;
    } // hInternet

CleanUp:

    // release our COM interfaces
    SAFE_RELEASE(pIWICStream);
    SAFE_RELEASE(pIWICDecoder);
    SAFE_RELEASE(pIWICBitmapFrameDecode);
    SAFE_RELEASE(pIWICConvertedFrame);

    // delete the bytes copied from those read from the Internet
    delete[] pBuf;
    pBuf = NULL;

    // close the Internet handles    
    if (hInternet)
    {
        InternetCloseHandle(hInternet);
    }

    return SUCCEEDED(hr) ? 0 : -1;
}
