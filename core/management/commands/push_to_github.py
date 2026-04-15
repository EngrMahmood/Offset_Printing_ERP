from django.core.management.base import BaseCommand
import subprocess
import os

class Command(BaseCommand):
    help = 'Push current changes to GitHub AI branch'

    def handle(self, *args, **options):
        try:
            # Check current branch
            result = subprocess.run(['git', 'branch', '--show-current'], 
                                  capture_output=True, text=True, cwd=os.getcwd())
            current_branch = result.stdout.strip()
            self.stdout.write(f'Current branch: {current_branch}')

            # If not on AI branch, switch to it
            if current_branch != 'AI':
                self.stdout.write('Switching to AI branch...')
                subprocess.run(['git', 'checkout', 'AI'], cwd=os.getcwd())
                self.stdout.write('Switched to AI branch')

            # Add all changes
            self.stdout.write('Adding changes...')
            subprocess.run(['git', 'add', '.'], cwd=os.getcwd())

            # Commit changes
            commit_message = 'feat: Add OEE calculations with impressions tracking, bulk upload Excel support, and dynamic templates'
            self.stdout.write(f'Committing with message: {commit_message}')
            result = subprocess.run(['git', 'commit', '-m', commit_message], 
                                  capture_output=True, text=True, cwd=os.getcwd())
            
            if result.returncode == 0:
                self.stdout.write('Changes committed successfully')
            else:
                self.stdout.write(f'Commit failed: {result.stderr}')

            # Push to remote
            self.stdout.write('Pushing to remote AI branch...')
            result = subprocess.run(['git', 'push', 'origin', 'AI'], 
                                  capture_output=True, text=True, cwd=os.getcwd())
            
            if result.returncode == 0:
                self.stdout.write('Successfully pushed to GitHub AI branch!')
            else:
                self.stdout.write(f'Push failed: {result.stderr}')

        except Exception as e:
            self.stdout.write(f'Error: {str(e)}')