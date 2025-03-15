#!/bin/bash

# Function to check and handle errors
handle_error() {
    if [ $? -ne 0 ]; then
        echo "Error: $1"
        exit 1
    fi
}

# Function to create HTML file
create_html_file() {
    local name="$1"
    local content="$2"
    local file_path="src/$name.html"
    echo "$content" > "$file_path"
    echo "Created $file_path"
}

# Check if clasp is installed
if ! command -v clasp &> /dev/null; then
    echo "Installing clasp..."
    npm install -g @google/clasp
    handle_error "Failed to install clasp"
fi

# Login to clasp if not already logged in
echo "Logging into clasp..."
clasp login
handle_error "Failed to login to clasp"

# Create new Google Sheet
echo "Creating new Google Sheet..."
SHEET_URL=$(curl -X POST \
    -H "Authorization: Bearer $(gcloud auth print-access-token)" \
    -H "Content-Type: application/json" \
    -d '{"mimeType": "application/vnd.google-apps.spreadsheet", "name": "Leave Management System"}' \
    "https://www.googleapis.com/drive/v3/files" | grep -o '"id":[^,}]*' | cut -d'"' -f4)

if [ -z "$SHEET_URL" ]; then
    echo "Creating sheet manually..."
    echo "Please create a new Google Sheet at https://sheets.new"
    echo "Copy the Sheet ID from the URL (the long string between /d/ and /edit)"
    echo ""
    read -p "Enter the Google Sheet ID: " SHEET_URL
fi

if [ -z "$SHEET_URL" ]; then
    echo "Error: Sheet ID is required"
    exit 1
fi

# Clean up existing configuration
echo "Cleaning up existing configuration..."
rm -f .clasp.json
handle_error "Failed to remove .clasp.json"

# Create new Apps Script project bound to the Sheet
echo "Creating new container-bound script..."
PROJECT_INFO=$(clasp create --type sheets --title "Leave Management System" --parentId "$SHEET_URL" --rootDir src 2>&1)
handle_error "Failed to create new project"

# Extract script ID from the output
SCRIPT_ID=$(echo "$PROJECT_INFO" | grep -o 'Created new Google Apps Script File:.*' | cut -d' ' -f6)

if [ -z "$SCRIPT_ID" ]; then
    echo "Error: Could not extract Script ID"
    exit 1
fi

# Create new .clasp.json
echo "{
  \"scriptId\": \"$SCRIPT_ID\",
  \"rootDir\": \"src\"
}" > .clasp.json
handle_error "Failed to create .clasp.json"

# Ensure appsscript.json is in src directory
if [ ! -f "src/appsscript.json" ]; then
    if [ -f "appsscript.json" ]; then
        mv appsscript.json src/
        handle_error "Failed to move appsscript.json"
    else
        echo "Error: appsscript.json not found"
        exit 1
    fi
fi

# Push all files to the project
echo "Pushing files to Apps Script..."
clasp push -f
handle_error "Failed to push files"

# Deploy as web app
echo "Deploying as web app..."
DEPLOYMENT_ID=$(clasp deploy --description "Leave Management System" | grep -o '"deploymentId": "[^"]*"' | cut -d'"' -f4)
handle_error "Failed to deploy web app"

echo ""
echo "Deployment successful!"
echo ""
echo "The system has been set up with:"
echo "1. Container-bound Apps Script project"
echo "2. All necessary code files"
echo "3. HTML templates for UI components"
echo ""
echo "Next steps:"
echo "1. Open your Google Sheet at: https://docs.google.com/spreadsheets/d/$SHEET_URL"
echo "2. Go to Extensions > Apps Script"
echo "3. Click on 'Run' > 'Run function' > 'initializeSystem'"
echo "4. Grant the necessary permissions when prompted"
echo ""
echo "The system will automatically:"
echo "- Create all required sheets"
echo "- Set up role management"
echo "- Configure interfaces"
echo "- Send access emails"
echo ""
echo "Google Sheet URL: https://docs.google.com/spreadsheets/d/$SHEET_URL"
echo "Script ID: $SCRIPT_ID"
echo "Deployment ID: $DEPLOYMENT_ID"
