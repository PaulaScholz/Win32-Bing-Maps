# Download and display Bing Maps with pure Win32
## Windows Developer Incubation and Learning - Paula Scholz

This project is an update of an earlier [version](https://www.codeproject.com/articles/170920/download-a-google-map-with-win32-c-and-wininet) published in 2011 for use on WindowsCE devices.  In this revision for the Windows Desktop, we use the [Bing Maps Platform](https://www.microsoft.com/en-us/maps) to download and display a static map image in a pure [Win32 API](https://docs.microsoft.com/en-us/windows/win32/) application.  Using [WinInet](https://docs.microsoft.com/en-us/windows/win32/wininet/about-wininet), [Vectors](https://docs.microsoft.com/en-us/cpp/standard-library/vector-class?view=vs-2019) from the [Standard Template Library](https://docs.microsoft.com/en-us/cpp/standard-library/cpp-standard-library-reference?view=vs-2019), the [Windows Imaging Component](https://docs.microsoft.com/en-us/windows/win32/wic/-wic-about-windows-imaging-codec), and [Win32 GDI](https://docs.microsoft.com/en-us/windows/win32/gdi/windows-gdi), we open an Internet connection to download and display a [static Bing Map](https://docs.microsoft.com/en-us/bingmaps/rest-services/imagery/get-a-static-map) on desktop Windows.

![GraphicsTestWin32 Application](ReadmeImages\SeattleMap.png)

To use the [Bing Maps APIs](https://docs.microsoft.com/en-us/bingmaps/rest-services/), you must have a [Bing Maps Key](https://docs.microsoft.com/en-us/bingmaps/getting-started/bing-maps-dev-center-help/getting-a-bing-maps-key).  These are easy to obtain and at the time of writing (April 2020) are free for development use.


