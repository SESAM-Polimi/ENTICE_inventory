"""Install pinned, explicitly patched MARIO/CVXLAB sources in a fresh environment.

Run with Python 3.13 (tested with 3.13.14 on macOS arm64). Optional local Git
repositories provide immutable archives without fetching sources; working-tree
changes are never copied. Dependency installation still needs PyPI or a populated
pip cache. The existing Python installation and source repositories are untouched.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
import tarfile
import venv

ROOT = Path(__file__).resolve().parents[1]


def run(arguments, *, cwd=None, env=None):
    print("Running:", " ".join(map(str, arguments)), flush=True)
    subprocess.run(list(map(str, arguments)), cwd=cwd, env=env, check=True)


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def prepare_source(name, definition, destination, local_repository):
    destination.mkdir(parents=True)
    revision = definition["revision"]
    if local_repository:
        archive = subprocess.check_output(["git", "-C", str(local_repository), "archive", revision])
        with tarfile.open(fileobj=io.BytesIO(archive)) as contents:
            contents.extractall(destination, filter="data")
        run(["git", "init", "--quiet", destination])
    else:
        run(["git", "init", "--quiet", destination])
        run(["git", "remote", "add", "origin", definition["repository"]], cwd=destination)
        run(["git", "fetch", "--depth", "1", "origin", revision], cwd=destination)
        fetched = subprocess.check_output(["git", "rev-parse", "FETCH_HEAD"], cwd=destination, text=True).strip()
        if fetched != revision:
            raise ValueError(f"Unexpected {name} revision: {fetched}")
        run(["git", "checkout", "--detach", "FETCH_HEAD"], cwd=destination)
    patch = ROOT / definition["patch"]
    if sha256(patch) != definition["patch_sha256"]:
        raise ValueError(f"Patch checksum mismatch for {name}")
    run(["git", "apply", "--check", patch], cwd=destination)
    run(["git", "apply", patch], cwd=destination)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--directory", type=Path, default=ROOT / "build/producer")
    parser.add_argument("--mario-repository", type=Path)
    parser.add_argument("--cvxlab-repository", type=Path)
    parser.add_argument("--cache-dir", type=Path)
    args = parser.parse_args()
    if sys.version_info[:2] != (3, 13):
        parser.error("Use Python 3.13 for this tested dependency lock.")
    destination = args.directory.resolve()
    if destination.exists():
        parser.error("Choose a new --directory; existing environments are not overwritten.")
    configuration = ROOT / "config/producer-environment.json"
    config = json.loads(configuration.read_text())
    destination.mkdir(parents=True)
    evidence = {"status": "installing", "python": platform.python_version(), "platform": platform.platform(),
                "configuration_sha256": sha256(configuration), "packages": config["packages"],
                "dependency_lock_sha256": sha256(ROOT / config["dependencies"])}
    try:
        for name, definition in config["packages"].items():
            prepare_source(name, definition, destination / "sources" / name, getattr(args, name + "_repository"))
        environment = destination / ".venv"
        venv.EnvBuilder(with_pip=True, system_site_packages=False).create(environment)
        python = environment / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
        env = os.environ.copy()
        env["PIP_CACHE_DIR"] = str((args.cache_dir or destination / "pip-cache").resolve())
        env["MPLCONFIGDIR"] = str(destination / "matplotlib-cache")
        run([python, "-m", "pip", "install", "-r", ROOT / config["dependencies"]], env=env)
        # Dependencies and build backend are pinned above. Disable build isolation
        # to avoid obtaining an unrecorded setuptools version during wheel creation.
        run([python, "-m", "pip", "install", "--no-build-isolation", "--no-deps",
             destination / "sources/mario", destination / "sources/cvxlab", ROOT], env=env)
        run([python, "-m", "pip", "check"], env=env)
        installed = subprocess.check_output([python, "-m", "pip", "list", "--format=json"], text=True, env=env)
        evidence.update(status="installed", installed_packages=json.loads(installed), python_executable=str(python))
    except Exception as exc:
        evidence.update(status="failed", error=f"{type(exc).__name__}: {exc}")
        raise
    finally:
        (destination / "environment.json").write_text(json.dumps(evidence, indent=2) + "\n")
    print(f"Environment ready: {python}")


if __name__ == "__main__":
    main()
