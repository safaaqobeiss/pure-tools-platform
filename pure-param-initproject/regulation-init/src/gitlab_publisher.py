"""
Publishes generated local files to GitLab repository projects.
"""
import base64
import logging
import urllib.parse
from pathlib import Path

import requests

from config import GitLabConfig

logger = logging.getLogger(__name__)


class GitLabPublisher:
    """Uploads local directories to GitLab projects using the Repository Files API."""

    def __init__(self, gitlab: GitLabConfig):
        self.gitlab = gitlab
        self._session = requests.Session()
        if gitlab.private_token:
            self._session.headers["PRIVATE-TOKEN"] = gitlab.private_token

    def publish_directory(
        self,
        project_path_with_namespace: str,
        local_dir: Path,
        branch: str = "main",
        commit_prefix: str = "chore: init files",
    ):
        """Create or update all files from *local_dir* into a GitLab project."""
        if not self.gitlab.private_token:
            raise ValueError("GitLab token is required for publishing. Use --token or GITLAB_TOKEN.")
        if not local_dir.exists():
            raise ValueError(f"Local directory not found: {local_dir}")

        project_id = self._encoded_project_id(project_path_with_namespace)

        # Ensure branch exists (create from default branch if needed)
        self._ensure_branch_exists(project_id, project_path_with_namespace, branch)

        files = [p for p in local_dir.rglob("*") if p.is_file()]
        if not files:
            logger.warning("No files found to publish in %s", local_dir)
            return

        logger.info(
            "Publishing %d file(s) to GitLab project '%s' (branch=%s)",
            len(files), project_path_with_namespace, branch,
        )

        for file_path in files:
            relative = file_path.relative_to(local_dir).as_posix()
            self._create_or_update_file(
                project_id=project_id,
                remote_file_path=relative,
                local_file=file_path,
                branch=branch,
                commit_message=f"{commit_prefix}: {relative}",
            )

    def _ensure_branch_exists(self, project_id: str, project_path: str, branch: str):
        """Check if branch exists; if not, create it from the default branch."""
        url = f"{self.gitlab.api_url}/projects/{project_id}/repository/branches/{urllib.parse.quote(branch, safe='')}"
        resp = self._session.get(url, timeout=30)
        if resp.status_code == 200:
            logger.info("Branche '%s' existe déjà dans %s", branch, project_path)
            return

        # Branch doesn't exist — find default branch
        proj_url = f"{self.gitlab.api_url}/projects/{project_id}"
        proj_resp = self._session.get(proj_url, timeout=30)
        default_branch = "main"
        if proj_resp.status_code == 200:
            proj_data = proj_resp.json() or {}
            default_branch = proj_data.get("default_branch") or "main"

        # If target branch == default branch, the repo is likely empty — that's ok
        if branch == default_branch:
            logger.info("Branche cible = branche par défaut '%s'", branch)
            return

        # Create branch from default
        create_url = f"{self.gitlab.api_url}/projects/{project_id}/repository/branches"
        payload = {"branch": branch, "ref": default_branch}
        create_resp = self._session.post(create_url, json=payload, timeout=30)
        if create_resp.status_code in (200, 201):
            logger.info("Branche '%s' créée depuis '%s' dans %s", branch, default_branch, project_path)
        elif create_resp.status_code == 400 and "already exists" in (create_resp.text or "").lower():
            logger.info("Branche '%s' existe déjà", branch)
        else:
            logger.warning(
                "Impossible de créer la branche '%s' (HTTP %s): %s",
                branch, create_resp.status_code, create_resp.text[:300],
            )

    def ensure_project_exists(
        self,
        project_path_with_namespace: str,
        username: str,
        create_if_missing: bool,
    ):
        """Ensure a project exists, optionally creating it in the given user's namespace."""
        project_id = self._encoded_project_id(project_path_with_namespace)
        url = f"{self.gitlab.api_url}/projects/{project_id}"
        resp = self._session.get(url, timeout=30)
        if resp.status_code == 200:
            logger.info("Projet GitLab trouvé: %s", project_path_with_namespace)
            return
        if resp.status_code == 404:
            logger.warning("Projet GitLab introuvable: %s", project_path_with_namespace)
        if resp.status_code != 404 or not create_if_missing:
            self._raise_gitlab_error(resp, f"Cannot access project '{project_path_with_namespace}'")

        # Try to create the project via API
        user = self._get_user(username)
        if not user:
            raise ValueError(f"Unable to find GitLab user '{username}' to create project.")

        namespace_path = project_path_with_namespace.rsplit("/", 1)[0]
        namespace_id = self._resolve_namespace_id(namespace_path, user)
        if not namespace_id:
            raise ValueError(
                f"Unable to resolve namespace id for '{namespace_path}'. "
                "Check your namespace path and token permissions."
            )

        project_name = project_path_with_namespace.rsplit("/", 1)[-1]
        create_url = f"{self.gitlab.api_url}/projects"
        payload = {
            "name": project_name,
            "path": project_name,
            "namespace_id": namespace_id,
            "initialize_with_readme": False,
            "visibility": "private",
        }
        create_resp = self._session.post(create_url, json=payload, timeout=30)
        if create_resp.status_code in (200, 201):
            logger.info("Created GitLab project: %s", project_path_with_namespace)
            return

        # Creation failed (e.g. 403) – ask user to create manually
        logger.warning(
            "Impossible de créer le projet via l'API (HTTP %s). "
            "Ton compte n'a peut-être pas la permission de créer des projets.",
            create_resp.status_code,
        )
        gitlab_base = self.gitlab.base_url.rstrip("/")
        print()
        print("=" * 60)
        print(f"Le projet '{project_path_with_namespace}' n'existe pas")
        print(f"et la création automatique a échoué (HTTP {create_resp.status_code}).")
        print()
        print("Crée-le manuellement sur GitLab :")
        print(f"  {gitlab_base}/projects/new#blank_project")
        print(f"  Nom du projet : {project_name}")
        print(f"  Namespace     : {namespace_path}")
        print("  Ne PAS cocher 'Initialize with a README'")
        print("=" * 60)

        while True:
            choice = input("Appuie sur Entrée une fois le projet créé (ou 'q' pour quitter): ").strip().lower()
            if choice == "q":
                raise RuntimeError(f"Projet '{project_path_with_namespace}' non créé. Abandon.")
            # Re-check if the project now exists
            check_resp = self._session.get(url, timeout=30)
            if check_resp.status_code == 200:
                logger.info("Projet GitLab trouvé après création manuelle: %s", project_path_with_namespace)
                return
            print(f"Projet toujours introuvable (HTTP {check_resp.status_code}). Réessaie.")

    def get_authenticated_username(self) -> str:
        """Return the GitLab username associated with the current private token."""
        if not self.gitlab.private_token:
            raise ValueError("GitLab token is required to resolve authenticated username.")
        url = f"{self.gitlab.api_url}/user"
        resp = self._session.get(url, timeout=30)
        if resp.status_code != 200:
            self._raise_gitlab_error(resp, "Failed to resolve authenticated GitLab user")
        data = resp.json() or {}
        username = str(data.get("username", "")).strip()
        if not username:
            raise RuntimeError("GitLab API did not return a username for the authenticated user.")
        return username

    def _get_user(self, username: str) -> dict | None:
        url = f"{self.gitlab.api_url}/users"
        resp = self._session.get(url, params={"username": username}, timeout=30)
        if resp.status_code != 200:
            self._raise_gitlab_error(resp, f"Failed to query GitLab user '{username}'")
        users = resp.json() or []
        return users[0] if users else None

    def _resolve_namespace_id(self, namespace_path: str, user: dict | None = None) -> int | None:
        """Resolve namespace id from namespace path, with fallback to user namespace_id."""
        # First try to resolve via the namespaces API (works for groups)
        url = f"{self.gitlab.api_url}/namespaces"
        resp = self._session.get(url, params={"search": namespace_path}, timeout=30)
        if resp.status_code == 200:
            namespaces = resp.json() or []
            for ns in namespaces:
                full_path = str(ns.get("full_path", "")).strip().lower()
                path = str(ns.get("path", "")).strip().lower()
                target = namespace_path.strip().lower()
                if full_path == target or path == target:
                    ns_id = ns.get("id")
                    if isinstance(ns_id, int):
                        return ns_id

        # Also try the groups API for nested group paths
        groups_url = f"{self.gitlab.api_url}/groups/{urllib.parse.quote(namespace_path, safe='')}"
        groups_resp = self._session.get(groups_url, timeout=30)
        if groups_resp.status_code == 200:
            group_data = groups_resp.json() or {}
            group_id = group_data.get("id")
            if isinstance(group_id, int):
                return group_id

        # Fallback to user namespace_id
        if user:
            user_namespace_id = user.get("namespace_id")
            if isinstance(user_namespace_id, int):
                return user_namespace_id

        return None

    def _create_or_update_file(
        self,
        project_id: str,
        remote_file_path: str,
        local_file: Path,
        branch: str,
        commit_message: str,
    ):
        encoded_file = urllib.parse.quote(remote_file_path, safe="")
        url = f"{self.gitlab.api_url}/projects/{project_id}/repository/files/{encoded_file}"

        content_b64 = base64.b64encode(local_file.read_bytes()).decode("ascii")
        payload = {
            "branch": branch,
            "start_branch": branch,
            "commit_message": commit_message,
            "content": content_b64,
            "encoding": "base64",
        }

        create_resp = self._session.post(url, json=payload, timeout=60)
        if create_resp.status_code in (200, 201):
            logger.debug("Created file: %s", remote_file_path)
            return

        if create_resp.status_code in (400, 409):
            update_resp = self._session.put(url, json=payload, timeout=60)
            if update_resp.status_code in (200, 201):
                logger.debug("Updated file: %s", remote_file_path)
                return
            if update_resp.status_code == 400 and "empty" in (update_resp.text or "").lower():
                logger.debug("Skipped unchanged file: %s", remote_file_path)
                return
            self._raise_gitlab_error(update_resp, f"Failed to update file '{remote_file_path}'")
        else:
            self._raise_gitlab_error(create_resp, f"Failed to create file '{remote_file_path}'")

    @staticmethod
    def _encoded_project_id(project_path_with_namespace: str) -> str:
        return urllib.parse.quote(project_path_with_namespace, safe="")

    @staticmethod
    def _raise_gitlab_error(resp: requests.Response, context: str):
        snippet = resp.text[:600] if resp.text else ""
        raise RuntimeError(f"{context}. HTTP {resp.status_code}. Response: {snippet}")