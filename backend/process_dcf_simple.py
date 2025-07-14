#!/usr/bin/env python3
"""
Simple DCF Model Processor - Uses existing infrastructure to process DCF uploads
"""

import asyncio
import os
import sys
from pathlib import Path

# Add tools directory to path
sys.path.append(str(Path(__file__).parent / "tools"))

from bulk_model_loader import BulkModelLoader

async def process_dcf_uploads():
    """Process DCF models in uploads folder using existing bulk loader"""
    
    # Get uploads directory (main spreadly/uploads folder)
    uploads_dir = Path(__file__).parent.parent / "uploads"
    
    print("🚀 DCF Model Processor")
    print("=" * 50)
    print(f"📁 Checking uploads directory: {uploads_dir}")
    
    if not uploads_dir.exists():
        uploads_dir.mkdir(parents=True, exist_ok=True)
        print(f"✅ Created uploads directory")
    
    # Check for Excel files
    excel_files = list(uploads_dir.glob("*.xlsx")) + list(uploads_dir.glob("*.xls"))
    
    print(f"📄 Found {len(excel_files)} Excel files")
    
    if not excel_files:
        print("💡 No Excel files found in uploads folder")
        print(f"💡 Upload your DCF models to: {uploads_dir}")
        return
    
    # List found files
    for i, file_path in enumerate(excel_files, 1):
        size_mb = file_path.stat().st_size / (1024 * 1024)
        print(f"  {i}. {file_path.name} ({size_mb:.1f} MB)")
    
    # Process with bulk loader
    print(f"\n🔄 Processing files with bulk loader...")
    
    loader = BulkModelLoader()
    
    # Show current stats
    stats_before = await loader.get_vector_store_stats()
    print(f"📊 Vector store before: {stats_before.get('total_models', 0)} models")
    
    # Process the uploads directory
    results = await loader.load_from_xlsx_directory(str(uploads_dir), auto_detect=True)
    
    # Show results
    print(f"\n📊 Processing Results:")
    print(f"Total processed: {results['total_processed']}")
    print(f"Successful: {results['successful']}")
    print(f"Failed: {results['failed']}")
    
    if results['loaded_models']:
        print(f"\n✅ Successfully Loaded DCF Models:")
        for model in results['loaded_models']:
            print(f"  • {model['file']} -> {model['model_id']}")
            print(f"    Type: {model['type']}, Industry: {model['industry']}")
    
    if results['errors']:
        print(f"\n❌ Errors:")
        for error in results['errors']:
            print(f"  • {error}")
    
    # Show final stats
    stats_after = await loader.get_vector_store_stats()
    print(f"\n📊 Vector store after: {stats_after.get('total_models', 0)} models")
    
    # Move processed files to processed folder
    if results['successful'] > 0:
        processed_dir = uploads_dir / "processed"
        processed_dir.mkdir(exist_ok=True)
        
        for excel_file in excel_files:
            new_path = processed_dir / excel_file.name
            if not new_path.exists():
                excel_file.rename(new_path)
                print(f"📁 Moved {excel_file.name} to processed/")
    
    return results

async def show_status():
    """Show current status"""
    uploads_dir = Path(__file__).parent.parent / "uploads"
    
    print("📊 DCF Upload Status")
    print("=" * 30)
    
    # Check uploads folder
    if uploads_dir.exists():
        excel_files = list(uploads_dir.glob("*.xlsx")) + list(uploads_dir.glob("*.xls"))
        processed_dir = uploads_dir / "processed"
        processed_files = []
        if processed_dir.exists():
            processed_files = list(processed_dir.glob("*.xlsx")) + list(processed_dir.glob("*.xls"))
        
        print(f"📁 Upload Directory: {uploads_dir}")
        print(f"📄 Pending files: {len(excel_files)}")
        if excel_files:
            for f in excel_files:
                print(f"   • {f.name}")
        print(f"✅ Processed files: {len(processed_files)}")
        if processed_files:
            for f in processed_files[-3:]:  # Show last 3
                print(f"   • {f.name}")
    else:
        print(f"📁 Upload Directory: Not found")
    
    # Check vector store
    loader = BulkModelLoader()
    stats = await loader.get_vector_store_stats()
    print(f"\n🗃️  Vector Store: {stats.get('total_models', 0)} total models")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "status":
        asyncio.run(show_status())
    else:
        asyncio.run(process_dcf_uploads())