#!/usr/bin/env python3
"""
DCF Upload Manager - Easy management interface for DCF model uploads
"""

import asyncio
import json
import sys
import os
from pathlib import Path
from datetime import datetime

# Add parent directory to path to import app modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dcf_model_processor import DCFModelProcessor
from upload_watcher import UploadWatcher
from app.services.model_vector_store import get_vector_store

class DCFUploadManager:
    """
    Easy-to-use manager for DCF upload processing
    """
    
    def __init__(self):
        self.processor = DCFModelProcessor()
        self.upload_dir = Path(os.path.dirname(os.path.dirname(__file__))) / 'uploads'
        self.vector_store = get_vector_store()
    
    async def status(self):
        """Show current status of uploads and vector store"""
        print("📊 DCF Upload System Status")
        print("=" * 50)
        
        # Check uploads directory
        if self.upload_dir.exists():
            excel_files = list(self.upload_dir.glob("*.xlsx")) + list(self.upload_dir.glob("*.xls"))
            processed_dir = self.upload_dir / "processed"
            processed_files = []
            if processed_dir.exists():
                processed_files = list(processed_dir.glob("*.xlsx")) + list(processed_dir.glob("*.xls"))
            
            print(f"📁 Upload Directory: {self.upload_dir}")
            print(f"📄 Pending files: {len(excel_files)}")
            print(f"✅ Processed files: {len(processed_files)}")
        else:
            print(f"📁 Upload Directory: {self.upload_dir} (not found)")
            print("📄 Pending files: 0")
            print("✅ Processed files: 0")
        
        # Check vector store status
        if self.vector_store.is_available():
            stats = self.vector_store.get_stats()
            print(f"\n🗃️  Vector Store Status: Available")
            print(f"📚 Total models: {stats.get('total_models', 0)}")
            print(f"🎯 DCF models: {stats.get('dcf_models', 'Unknown')}")
        else:
            print(f"\n🗃️  Vector Store Status: Not Available")
        
        print()
    
    async def process_all(self):
        """Process all pending DCF models in uploads folder"""
        print("🔄 Processing All DCF Models")
        print("=" * 40)
        
        results = await self.processor.process_dcf_uploads(str(self.upload_dir))
        
        print(f"\n📊 Results:")
        print(f"Total files found: {results['total_found']}")
        print(f"DCF models identified: {results['dcf_models_processed']}")
        print(f"Successfully processed: {results['successful']}")
        print(f"Failed: {results['failed']}")
        
        if results['processed_models']:
            print(f"\n✅ Successfully Processed:")
            for model in results['processed_models']:
                print(f"  • {model['file']}")
                print(f"    Quality: {model['quality_score']:.2f}/5.0")
                print(f"    Industry: {model['industry']}")
                print(f"    Components: {', '.join(model['dcf_components'])}")
                print()
        
        if results['errors']:
            print(f"\n❌ Errors:")
            for error in results['errors']:
                print(f"  • {error}")
        
        return results
    
    async def watch_folder(self, interval: int = 30):
        """Start watching the uploads folder for new files"""
        print(f"👀 Starting Upload Watcher (checking every {interval}s)")
        print("Press Ctrl+C to stop")
        print("=" * 50)
        
        watcher = UploadWatcher(str(self.upload_dir), interval)
        
        # Process existing files first
        await watcher.process_existing_files()
        
        # Start watching
        await watcher.start_watching()
    
    async def analyze_file(self, filename: str):
        """Analyze a specific file to see if it's a DCF model"""
        file_path = self.upload_dir / filename
        
        if not file_path.exists():
            print(f"❌ File not found: {filename}")
            return
        
        print(f"🔍 Analyzing: {filename}")
        print("=" * 40)
        
        analysis = self.processor._analyze_dcf_model(str(file_path))
        
        print(f"DCF Model: {'✅ Yes' if analysis['is_dcf_model'] else '❌ No'}")
        print(f"Quality Score: {analysis['quality_score']:.2f}/5.0")
        print(f"Industry: {analysis['industry'].value if 'industry' in analysis else 'Unknown'}")
        print(f"Complexity: {analysis['complexity'].value if 'complexity' in analysis else 'Unknown'}")
        
        if analysis['components_found']:
            print(f"\nComponents Found:")
            for component in analysis['components_found']:
                print(f"  • {component}")
        
        if analysis.get('suggested_improvements'):
            print(f"\nSuggested Improvements:")
            for improvement in analysis['suggested_improvements']:
                print(f"  • {improvement}")
        
        return analysis
    
    async def list_files(self):
        """List all files in uploads directory"""
        print("📁 Files in Uploads Directory")
        print("=" * 40)
        
        if not self.upload_dir.exists():
            print("Upload directory not found")
            return
        
        excel_files = list(self.upload_dir.glob("*.xlsx")) + list(self.upload_dir.glob("*.xls"))
        
        if not excel_files:
            print("No Excel files found")
            return
        
        for i, file_path in enumerate(excel_files, 1):
            size_mb = file_path.stat().st_size / (1024 * 1024)
            modified = datetime.fromtimestamp(file_path.stat().st_mtime)
            print(f"{i:2d}. {file_path.name}")
            print(f"     Size: {size_mb:.1f} MB, Modified: {modified.strftime('%Y-%m-%d %H:%M')}")
    
    async def setup_folder(self):
        """Setup the uploads folder structure"""
        print("🔧 Setting up uploads folder structure")
        print("=" * 40)
        
        # Create main uploads directory
        self.upload_dir.mkdir(parents=True, exist_ok=True)
        print(f"✅ Created: {self.upload_dir}")
        
        # Create processed subdirectory
        processed_dir = self.upload_dir / "processed"
        processed_dir.mkdir(exist_ok=True)
        print(f"✅ Created: {processed_dir}")
        
        # Create README
        readme_content = """# DCF Models Upload Folder

## Usage:
1. Drop your DCF Excel models (.xlsx or .xls) into this folder
2. Run the processor: `python tools/manage_dcf_uploads.py process`
3. Or start auto-watching: `python tools/manage_dcf_uploads.py watch`

## File Processing:
- Files are automatically analyzed for DCF content
- Quality scores are calculated (0-5.0 scale)
- Successfully processed files are moved to ./processed/
- Processing logs are saved as processing_results.log

## Quality Scoring:
- 4.0-5.0: High quality, comprehensive DCF model
- 2.5-3.9: Medium quality, some DCF components
- 0-2.4: Basic model, may need improvements

## Supported Features:
- WACC/Cost of Capital detection
- Free Cash Flow projections
- Terminal Value calculations
- Assumptions sections
- Multi-year forecasts
- Industry classification
- Complexity assessment
"""
        
        readme_path = self.upload_dir / "README.md"
        with open(readme_path, 'w') as f:
            f.write(readme_content)
        print(f"✅ Created: {readme_path}")
        
        print(f"\n🎉 Setup complete! You can now upload DCF models to:")
        print(f"    {self.upload_dir}")

async def main():
    import argparse
    
    parser = argparse.ArgumentParser(description="DCF Upload Manager")
    parser.add_argument('command', choices=['status', 'process', 'watch', 'analyze', 'list', 'setup'],
                       help='Command to execute')
    parser.add_argument('--file', type=str, help='Filename for analyze command')
    parser.add_argument('--interval', type=int, default=30, help='Watch interval in seconds')
    
    args = parser.parse_args()
    
    manager = DCFUploadManager()
    
    if args.command == 'status':
        await manager.status()
    
    elif args.command == 'process':
        await manager.process_all()
    
    elif args.command == 'watch':
        await manager.watch_folder(args.interval)
    
    elif args.command == 'analyze':
        if not args.file:
            print("❌ Please specify --file for analyze command")
            return
        await manager.analyze_file(args.file)
    
    elif args.command == 'list':
        await manager.list_files()
    
    elif args.command == 'setup':
        await manager.setup_folder()

if __name__ == "__main__":
    asyncio.run(main())