#!/bin/bash

# DCF Upload Management Script
# Easy commands for managing DCF model uploads

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON_SCRIPT="$SCRIPT_DIR/manage_dcf_uploads.py"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Function to print colored output
print_color() {
    echo -e "${1}${2}${NC}"
}

# Function to show usage
show_usage() {
    print_color $BLUE "🚀 DCF Upload Manager"
    echo "================================"
    echo ""
    echo "Usage: ./dcf_upload.sh [command]"
    echo ""
    echo "Commands:"
    echo "  setup     - Setup uploads folder structure"
    echo "  status    - Show current status"
    echo "  process   - Process all pending DCF models"
    echo "  watch     - Start watching for new uploads (auto-process)"
    echo "  list      - List files in uploads folder"
    echo "  analyze   - Analyze a specific file"
    echo "  help      - Show this help message"
    echo ""
    echo "Examples:"
    echo "  ./dcf_upload.sh setup"
    echo "  ./dcf_upload.sh process"
    echo "  ./dcf_upload.sh watch"
    echo "  ./dcf_upload.sh analyze --file my_model.xlsx"
    echo ""
}

# Check if Python script exists
if [ ! -f "$PYTHON_SCRIPT" ]; then
    print_color $RED "❌ Error: manage_dcf_uploads.py not found"
    exit 1
fi

# Handle commands
case "$1" in
    "setup")
        print_color $BLUE "🔧 Setting up DCF uploads folder..."
        python3 "$PYTHON_SCRIPT" setup
        ;;
    "status")
        print_color $BLUE "📊 Checking DCF upload status..."
        python3 "$PYTHON_SCRIPT" status
        ;;
    "process")
        print_color $YELLOW "🔄 Processing all DCF models..."
        python3 "$PYTHON_SCRIPT" process
        ;;
    "watch")
        print_color $GREEN "👀 Starting upload watcher..."
        python3 "$PYTHON_SCRIPT" watch
        ;;
    "list")
        print_color $BLUE "📁 Listing files..."
        python3 "$PYTHON_SCRIPT" list
        ;;
    "analyze")
        if [ -z "$3" ]; then
            print_color $RED "❌ Please specify a filename: ./dcf_upload.sh analyze --file filename.xlsx"
            exit 1
        fi
        print_color $YELLOW "🔍 Analyzing file..."
        python3 "$PYTHON_SCRIPT" analyze "$2" "$3"
        ;;
    "help"|"--help"|"-h"|"")
        show_usage
        ;;
    *)
        print_color $RED "❌ Unknown command: $1"
        echo ""
        show_usage
        exit 1
        ;;
esac