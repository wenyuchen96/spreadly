#!/usr/bin/env python3
"""
Upload Watcher - Monitors uploads folder and automatically processes DCF models
"""

import asyncio
import time
import logging
from pathlib import Path
from typing import Set
import sys
import os
from datetime import datetime

# Add parent directory to path to import app modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dcf_model_processor import DCFModelProcessor

class UploadWatcher:
    """
    Watches the uploads folder and automatically processes new DCF models
    """
    
    def __init__(self, upload_directory: str = None, check_interval: int = 30):
        """
        Initialize the upload watcher
        
        Args:
            upload_directory: Path to monitor (defaults to ../uploads)
            check_interval: How often to check for new files (seconds)
        """
        if upload_directory is None:
            upload_directory = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'uploads')
        
        self.upload_directory = Path(upload_directory)
        self.check_interval = check_interval
        self.processor = DCFModelProcessor()
        self.processed_files: Set[str] = set()
        self.is_running = False
        
        # Ensure uploads directory exists
        self.upload_directory.mkdir(parents=True, exist_ok=True)
        
        # Setup logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('upload_watcher.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
        # Load previously processed files
        self._load_processed_files()
    
    def _load_processed_files(self):
        """Load list of previously processed files"""
        processed_log = self.upload_directory / "processed_files.log"
        if processed_log.exists():
            with open(processed_log, 'r') as f:
                self.processed_files = set(line.strip() for line in f if line.strip())
            self.logger.info(f"Loaded {len(self.processed_files)} previously processed files")
    
    def _save_processed_file(self, filename: str):
        """Save a file as processed"""
        self.processed_files.add(filename)
        processed_log = self.upload_directory / "processed_files.log"
        with open(processed_log, 'a') as f:
            f.write(f"{filename}\n")
    
    async def start_watching(self):
        """Start monitoring the uploads folder"""
        self.is_running = True
        self.logger.info(f"🔍 Starting upload watcher for: {self.upload_directory}")
        self.logger.info(f"⏰ Check interval: {self.check_interval} seconds")
        
        while self.is_running:
            try:
                await self._check_for_new_files()
                await asyncio.sleep(self.check_interval)
            except KeyboardInterrupt:
                self.logger.info("📛 Received interrupt signal, stopping watcher...")
                self.stop_watching()
            except Exception as e:
                self.logger.error(f"❌ Error in watcher loop: {str(e)}")
                await asyncio.sleep(self.check_interval)
    
    def stop_watching(self):
        """Stop the watcher"""
        self.is_running = False
        self.logger.info("🛑 Upload watcher stopped")
    
    async def _check_for_new_files(self):
        """Check for new Excel files and process them"""
        try:
            # Find Excel files
            excel_files = list(self.upload_directory.glob("*.xlsx")) + list(self.upload_directory.glob("*.xls"))
            
            new_files = []
            for excel_file in excel_files:
                if excel_file.name not in self.processed_files:
                    # Check if file is still being written (wait for stable size)
                    if self._is_file_stable(excel_file):
                        new_files.append(excel_file)
            
            if new_files:
                self.logger.info(f"📁 Found {len(new_files)} new files to process")
                await self._process_new_files(new_files)
            
        except Exception as e:
            self.logger.error(f"❌ Error checking for new files: {str(e)}")
    
    def _is_file_stable(self, file_path: Path, stability_time: int = 5) -> bool:
        """Check if file has stopped changing (finished uploading)"""
        try:
            # Get current file size
            current_size = file_path.stat().st_size
            
            # Wait a bit and check again
            time.sleep(stability_time)
            new_size = file_path.stat().st_size
            
            # File is stable if size hasn't changed
            return current_size == new_size
        except Exception:
            return False
    
    async def _process_new_files(self, new_files: list):
        """Process new Excel files"""
        for excel_file in new_files:
            try:
                self.logger.info(f"🔄 Processing new file: {excel_file.name}")
                
                # Analyze if it's a DCF model
                analysis = self.processor._analyze_dcf_model(str(excel_file))
                
                if analysis['is_dcf_model']:
                    self.logger.info(f"✅ Identified DCF model: {excel_file.name} (Quality: {analysis['quality_score']:.2f}/5.0)")
                    
                    # Create and add to vector store
                    model = await self.processor._create_dcf_model(excel_file, analysis)
                    success = await self.processor.vector_store.add_model(model)
                    
                    if success:
                        self.logger.info(f"📚 Added to RAG: {model.name}")
                        
                        # Move to processed folder
                        self.processor._move_processed_file(excel_file)
                        
                        # Mark as processed
                        self._save_processed_file(excel_file.name)
                        
                        # Log processing result
                        self._log_processing_result(excel_file.name, model, analysis)
                        
                    else:
                        self.logger.error(f"❌ Failed to add to vector store: {excel_file.name}")
                else:
                    self.logger.info(f"⏭️  Not a DCF model: {excel_file.name}")
                    # Mark as processed so we don't keep checking it
                    self._save_processed_file(excel_file.name)
                    
            except Exception as e:
                self.logger.error(f"❌ Error processing {excel_file.name}: {str(e)}")
    
    def _log_processing_result(self, filename: str, model, analysis: dict):
        """Log detailed processing results"""
        result_log = self.upload_directory / "processing_results.log"
        
        result_entry = {
            "timestamp": datetime.now().isoformat(),
            "filename": filename,
            "model_id": model.id,
            "model_name": model.name,
            "industry": model.industry.value,
            "complexity": model.complexity.value,
            "quality_score": analysis['quality_score'],
            "components_found": analysis['components_found'],
            "tags": model.tags
        }
        
        with open(result_log, 'a') as f:
            f.write(f"{result_entry}\n")
    
    async def process_existing_files(self):
        """One-time processing of existing files in uploads folder"""
        self.logger.info("🔄 Processing existing files in uploads folder...")
        results = await self.processor.process_dcf_uploads(str(self.upload_directory))
        
        # Mark all processed files
        for model_info in results['processed_models']:
            self._save_processed_file(model_info['file'])
        
        self.logger.info(f"✅ Initial processing complete: {results['successful']} DCF models added")
        return results

# CLI interface
async def main():
    import argparse
    
    parser = argparse.ArgumentParser(description="DCF Upload Watcher")
    parser.add_argument('--upload-dir', type=str, help='Upload directory to monitor')
    parser.add_argument('--interval', type=int, default=30, help='Check interval in seconds')
    parser.add_argument('--process-existing', action='store_true', help='Process existing files first')
    parser.add_argument('--watch', action='store_true', help='Start continuous watching')
    
    args = parser.parse_args()
    
    watcher = UploadWatcher(
        upload_directory=args.upload_dir,
        check_interval=args.interval
    )
    
    print("🚀 DCF Upload Watcher")
    print("=" * 50)
    print(f"📁 Monitoring: {watcher.upload_directory}")
    print(f"⏰ Check interval: {args.interval} seconds")
    print()
    
    try:
        if args.process_existing:
            print("🔄 Processing existing files...")
            results = await watcher.process_existing_files()
            print(f"✅ Processed {results['dcf_models_processed']} DCF models")
            print()
        
        if args.watch:
            print("👀 Starting continuous monitoring...")
            print("Press Ctrl+C to stop")
            await watcher.start_watching()
        else:
            print("💡 Use --watch to start continuous monitoring")
            print("💡 Use --process-existing to process files already in folder")
            
    except KeyboardInterrupt:
        print("\n📛 Stopping watcher...")
        watcher.stop_watching()

if __name__ == "__main__":
    asyncio.run(main())