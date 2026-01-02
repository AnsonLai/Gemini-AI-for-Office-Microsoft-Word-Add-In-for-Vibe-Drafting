@echo off
echo Starting AI Word Plugin development server...
echo This will run in the background and serve the add-in files.
echo You can manually open Word to use the add-in.
echo.
echo To stop the server, close this window or press Ctrl+C
echo.

REM Start webpack dev server in development mode
REM The server will run on https://localhost:3000 as configured
start "AI Word Plugin Dev Server" npm run dev-server

echo.
echo Dev server started. The add-in is now available at https://localhost:3000
echo You can open Word manually to use the add-in.
echo.
echo Press any key to exit this script (the dev server will continue running)...
pause >nul