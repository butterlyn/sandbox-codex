"""Download the pinned AEMO originals and verify their SHA-256 digests."""
from concurrent.futures import ThreadPoolExecutor
import hashlib
import json
from pathlib import Path
import urllib.request

ROOT = Path(__file__).resolve().parents[1]


def download(source):
    destination = ROOT / "data/sources" / source["filename"]
    if destination.exists() and hashlib.sha256(destination.read_bytes()).hexdigest() == source["sha256"]:
        print(f"Verified {destination.name}")
        return
    temporary = destination.with_suffix(destination.suffix + ".part")
    request = urllib.request.Request(source["url"], headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=180) as response, temporary.open("wb") as output:
        while chunk := response.read(1024 * 1024):
            output.write(chunk)
    digest = hashlib.sha256(temporary.read_bytes()).hexdigest()
    if digest != source["sha256"]:
        temporary.unlink()
        raise ValueError(f"AEMO source changed: {destination.name}; expected {source['sha256']}, received {digest}")
    temporary.replace(destination)
    print(f"Downloaded and verified {destination.name}")


if __name__ == "__main__":
    manifest = json.loads((ROOT / "data/sources/sources.json").read_text())
    with ThreadPoolExecutor(max_workers=4) as executor:
        list(executor.map(download, manifest))
