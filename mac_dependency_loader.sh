#!/bin/bash
# This script installs packages needed for Rollout Automation bot and must be run with sudo privileges.

# Check if Python is installed
PYTHON_HOME=$(which python3 || which python)

if [ -x "$PYTHON_HOME" ]; then
    echo "Python location is in: $PYTHON_HOME"
    echo
else
    echo "Python does not exist in the system, please contact ITBar."
    echo
    exit 1
fi

# Common Packages - Webtool
list="gspread python-dotenv selenium webdriver-manager google-api-python-client google-auth oauth2client requests"

# MTPOS Specific
list="$list PyQt5 pywinauto"

# Track missing packages
notInstalled=()

# Check and collect missing packages
for pkg in $list; do
    if ! python3 -m pip show "$pkg" > /dev/null 2>&1; then
        echo "$pkg is NOT installed."
        notInstalled+=("$pkg")
    else
        echo "$pkg is installed. Removing from list."
    fi
done

# Display remaining packages
echo
echo "Remaining packages (to be installed):"
echo "${notInstalled[@]}"
echo

# Install missing packages
if [ ${#notInstalled[@]} -ne 0 ]; then
    echo "Installing missing packages..."
    python3 -m pip install "${notInstalled[@]}"
    echo "Installation Complete"
else
    echo "All packages are already installed."
fi

read -p "Press any key to continue..." -n1 -s
echo
