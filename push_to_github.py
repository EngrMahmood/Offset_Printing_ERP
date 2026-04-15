#!/usr/bin/env python3
import subprocess
import os

def run_git_command(command, cwd=None):
    """Run a git command and return the result"""
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True, cwd=cwd or os.getcwd())
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except Exception as e:
        return 1, "", str(e)

def main():
    print("Pushing changes to GitHub AI branch...")

    # Check current branch
    code, branch, error = run_git_command("git branch --show-current")
    if code != 0:
        print(f"Error checking branch: {error}")
        return

    print(f"Current branch: {branch}")

    # Switch to AI branch if not already on it
    if branch != "AI":
        print("Switching to AI branch...")
        code, output, error = run_git_command("git checkout AI")
        if code != 0:
            print(f"Error switching to AI branch: {error}")
            return
        print("Switched to AI branch")

    # Add all changes
    print("Adding changes...")
    code, output, error = run_git_command("git add .")
    if code != 0:
        print(f"Error adding changes: {error}")
        return
    print("Changes added")

    # Commit changes
    commit_message = "feat: Add OEE calculations with impressions tracking, bulk upload Excel support, and dynamic templates"
    print(f"Committing with message: {commit_message}")
    code, output, error = run_git_command(f'git commit -m "{commit_message}"')
    if code != 0:
        if "nothing to commit" in error:
            print("No changes to commit")
        else:
            print(f"Error committing: {error}")
            return
    else:
        print("Changes committed successfully")

    # Push to remote
    print("Pushing to remote AI branch...")
    code, output, error = run_git_command("git push origin AI")
    if code != 0:
        print(f"Error pushing: {error}")
        return

    print("Successfully pushed to GitHub AI branch!")

if __name__ == "__main__":
    main()