"""Reproduce the parallel chart build, then validate and assemble the single HTML."""
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
JOBS = [f"build_group_{i}.py" for i in range(1, 6)] + [
    "build_supplemental.py", "build_supplemental_network.py",
    "build_supplemental_events.py", "build_supplemental_duration.py",
]


def run(script):
    result = subprocess.run([sys.executable, str(ROOT / "scripts" / script)], cwd=ROOT,
                            text=True, capture_output=True)
    if result.returncode:
        raise RuntimeError(f"{script} failed:\n{result.stdout}\n{result.stderr}")
    return script, result.stdout.strip()


if __name__ == "__main__":
    for script in ("download_sources.py", "extract_inventory.py"):
        name, result = run(script)
        print(f"{name}: {result}", flush=True)
    with ThreadPoolExecutor(max_workers=6) as executor:
        futures = [executor.submit(run, job) for job in JOBS]
        for future in as_completed(futures):
            name, result = future.result()
            print(f"Completed {name}\n{result}", flush=True)
    for script in ("validate.py", "build.py"):
        name, result = run(script)
        print(f"{name}: {result}", flush=True)
