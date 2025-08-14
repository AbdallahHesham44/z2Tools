import base64
import requests
from github import Github
import io

def download_file_from_github(token, repo_name, path):
    g = Github(token)
    repo = g.get_repo(repo_name)
    contents = repo.get_contents(path)
    return contents.decoded_content

def upload_file_to_github(token, repo_name, path, file_bytes, commit_message="Update file"):
    g = Github(token)
    repo = g.get_repo(repo_name)
    try:
        contents = repo.get_contents(path)
        repo.update_file(contents.path, commit_message, file_bytes, contents.sha)
    except:
        repo.create_file(path, commit_message, file_bytes)
