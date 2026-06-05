"""Run the PLM HQ attachment download flow for one material number.

Run from the repository root:
    python web_app2/plm/try_hq_search.py --username YOUR_ID --hqpn HQ111120B1009
"""

from __future__ import annotations

import argparse
import getpass
from pathlib import Path

from playwright.sync_api import sync_playwright

try:
    from .automation import run_hq_attachment_download
except ImportError:
    from automation import run_hq_attachment_download


def _output_dir() -> Path:
    path = Path("web_app2") / "outputs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def run(username: str, password: str, hqpn: str) -> Path:
    with sync_playwright() as playwright:
        return run_hq_attachment_download(
            playwright,
            username=username,
            password=password,
            hqpn=hqpn,
            output_dir=_output_dir(),
            headless=False,
            log=lambda message: print(message, flush=True),
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="Try PLM HQ attachment download without coordinate clicks.")
    parser.add_argument("--username", required=True, help="EIP/PLM username")
    parser.add_argument("--hqpn", required=True, help="HQ material number")
    parser.add_argument("--password", help="EIP/PLM password; prompted if omitted")
    args = parser.parse_args()

    password = args.password or getpass.getpass("Password: ")
    output_path = run(args.username, password, args.hqpn)
    print(f"Download complete: {output_path}")


if __name__ == "__main__":
    main()
