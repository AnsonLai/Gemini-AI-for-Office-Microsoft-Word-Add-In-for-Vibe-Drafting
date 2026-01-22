@echo off
echo Resetting and installing Office Add-in development certificates...
echo This will help fix "The dev server is not running on port 3000" errors.
echo.

echo Step 1: Uninstalling existing certificates...
npx office-addin-dev-certs uninstall

echo.
echo Step 2: Installing new certificates...
npx office-addin-dev-certs install --static src/taskpane/taskpane.html

echo.
echo Step 3: Verifying certificates...
npx office-addin-dev-certs verify

echo.
echo Certificate fix complete. Please try running "npm start" again.
pause
