import os
import shutil
import subprocess
import tempfile
import json
import re
from textwrap import dedent
from typing import Dict, List, Any

# Global persistent temp directory for reuse
_PERSISTENT_TEMP_DIR = None
_TSC_PATH = None

def parse_tsc_errors(stdout: str) -> List[Dict[str, Any]]:
    """Parse TypeScript compiler errors into structured format."""
    errors = []
    # Pattern: filename(line,col): error TSxxxx: message
    pattern = r'(.+?)\((\d+),(\d+)\):\s+(error|warning)\s+TS(\d+):\s+(.+)'

    for match in re.finditer(pattern, stdout):
        filename, line, col, severity, code, message = match.groups()
        errors.append({
            'file': os.path.basename(filename),
            'line': int(line),
            'column': int(col),
            'severity': severity,
            'code': f'TS{code}',
            'message': message.strip()
        })

    return errors


def format_error_summary(errors: List[Dict[str, Any]]) -> str:
    """Format errors into a readable summary."""
    if not errors:
        return "No errors found"

    summary = []
    for i, err in enumerate(errors, 1):
        summary.append(
            f"  [{i}] Line {err['line']}:{err['column']} - "
            f"{err['severity'].upper()} {err['code']}\n"
            f"      {err['message']}"
        )

    return "\n".join(summary)


def setup_persistent_env(use_preview: bool = True) -> tuple:
    """
    Setup a persistent TypeScript environment that can be reused.
    Returns (temp_dir, tsc_path)
    """
    global _PERSISTENT_TEMP_DIR, _TSC_PATH

    # If already setup, return cached paths
    if _PERSISTENT_TEMP_DIR and _TSC_PATH and os.path.exists(_PERSISTENT_TEMP_DIR):
        return _PERSISTENT_TEMP_DIR, _TSC_PATH

    npm = "npm.cmd" if os.name == "nt" else "npm"
    if shutil.which(npm) is None:
        raise RuntimeError("npm not found on PATH. Install Node.js or add npm to PATH.")

    # Create persistent temp directory
    td = tempfile.mkdtemp(prefix="ojs_tscheck_persistent_")
    print(f"[SETUP] Creating persistent TypeScript environment in: {td}")

    # 1) package.json
    p = subprocess.run([npm, "init", "-y"], cwd=td, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError(f"npm init failed:\n{p.stdout}\n{p.stderr}")

    # 2) Install dependencies
    pkgs = ["typescript", "@types/office-js"]
    if use_preview:
        pkgs.append("@types/office-js-preview")

    print(f"[SETUP] Installing packages: {', '.join(pkgs)}")
    p = subprocess.run([npm, "i", "-D", "--silent", "--no-audit", "--no-fund", *pkgs],
                       cwd=td, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError(f"npm install failed:\n{p.stdout}\n{p.stderr}")

    # 3) Create tsconfig.json
    tsconfig = {
        "compilerOptions": {
            "target": "ES2018",
            "module": "ESNext",
            "lib": ["ES2018", "DOM"],
            "strict": False,
            "moduleResolution": "Node",
            "typeRoots": [os.path.join(td, "node_modules", "@types")],
            "types": ["office-js"] + (["office-js-preview"] if use_preview else []),
            "skipLibCheck": True
        }
    }
    with open(os.path.join(td, "tsconfig.json"), "w", encoding="utf-8") as f:
        json.dump(tsconfig, f, indent=2)

    # 4) Get tsc path
    tsc = os.path.join(td, "node_modules", ".bin", "tsc")
    if os.name == "nt":
        if os.path.exists(tsc + ".cmd"):
            tsc = tsc + ".cmd"

    # Cache the paths
    _PERSISTENT_TEMP_DIR = td
    _TSC_PATH = tsc

    print(f"[SETUP] Environment ready! tsc at: {tsc}\n")

    return td, tsc


def check_officejs_ts_fast(code_ts: str, use_preview: bool = True) -> dict:
    """
    Fast type-check using a persistent TypeScript environment.
    Only writes the snippet file and runs tsc, no npm install overhead.
    """
    # Setup persistent environment (only runs once)
    td, tsc = setup_persistent_env(use_preview)

    # Write snippet to a file in the persistent directory
    ts_path = os.path.join(td, "snippet.ts")
    with open(ts_path, "w", encoding="utf-8") as f:
        f.write(dedent(code_ts).strip() + "\n")

    # Update tsconfig to include this snippet
    tsconfig_path = os.path.join(td, "tsconfig.json")
    with open(tsconfig_path, "r", encoding="utf-8") as f:
        tsconfig = json.load(f)
    tsconfig["include"] = ["snippet.ts"]
    with open(tsconfig_path, "w", encoding="utf-8") as f:
        json.dump(tsconfig, f, indent=2)

    # Run tsc
    cmd = [tsc, "--noEmit"]
    proc = subprocess.run(cmd, cwd=td, capture_output=True, text=True)

    # Parse errors
    errors = parse_tsc_errors(proc.stdout)
    error_summary = format_error_summary(errors)

    # Determine success/failure
    success = proc.returncode == 0
    status = "PASS" if success else "FAIL"

    return {
        # Status
        "success": success,
        "status": status,
        "returncode": proc.returncode,

        # Errors
        "errors": errors,
        "error_count": len(errors),
        "error_summary": error_summary,

        # Raw output
        "stdout": proc.stdout,
        "stderr": proc.stderr,

        # Metadata
        "temp_dir": td,
        "ts_file": ts_path,
        "tsconfig": tsconfig_path,
        "cmd": " ".join(cmd),
        "use_preview": use_preview
    }


def cleanup_persistent_env():
    """Clean up the persistent environment when done."""
    global _PERSISTENT_TEMP_DIR, _TSC_PATH

    if _PERSISTENT_TEMP_DIR and os.path.exists(_PERSISTENT_TEMP_DIR):
        print(f"\n[CLEANUP] Removing temp directory: {_PERSISTENT_TEMP_DIR}")
        try:
            shutil.rmtree(_PERSISTENT_TEMP_DIR)
        except Exception as e:
            print(f"[CLEANUP] Warning: Failed to remove temp directory: {e}")

    _PERSISTENT_TEMP_DIR = None
    _TSC_PATH = None


# Keep the original function for backward compatibility
def check_officejs_ts(code_ts: str, use_preview: bool = True) -> dict:
    """
    Original function - creates new temp env each time (SLOW).
    Use check_officejs_ts_fast() instead for better performance.
    """
    npm = "npm.cmd" if os.name == "nt" else "npm"
    if shutil.which(npm) is None:
        raise RuntimeError("npm not found on PATH. Install Node.js or add npm to PATH.")

    # 1) temp project
    td = tempfile.mkdtemp(prefix="ojs_tscheck_")
    ts_path = os.path.join(td, "snippet.ts")
    with open(ts_path, "w", encoding="utf-8") as f:
        f.write(dedent(code_ts).strip() + "\n")

    # 2) package.json
    p = subprocess.run([npm, "init", "-y"], cwd=td, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError(f"npm init failed:\n{p.stdout}\n{p.stderr}")

    # 3) deps
    pkgs = ["typescript", "@types/office-js"]
    if use_preview:
        pkgs.append("@types/office-js-preview")
    p = subprocess.run([npm, "i", "-D", "--silent", "--no-audit", "--no-fund", *pkgs],
                       cwd=td, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError(f"npm install failed:\n{p.stdout}\n{p.stderr}")

    # 4) tsconfig.json
    tsconfig = {
        "compilerOptions": {
            "target": "ES2018",
            "module": "ESNext",
            "lib": ["ES2018", "DOM"],
            "strict": False,
            "moduleResolution": "Node",
            "typeRoots": [os.path.join(td, "node_modules", "@types")],
            "types": ["office-js"] + (["office-js-preview"] if use_preview else []),
            "skipLibCheck": True
        },
        "include": ["snippet.ts"]
    }
    with open(os.path.join(td, "tsconfig.json"), "w", encoding="utf-8") as f:
        json.dump(tsconfig, f, indent=2)

    # 5) run local tsc
    tsc = os.path.join(td, "node_modules", ".bin", "tsc")
    if os.name == "nt":
        if os.path.exists(tsc + ".cmd"):
            tsc = tsc + ".cmd"
    cmd = [tsc, "--noEmit"]
    proc = subprocess.run(cmd, cwd=td, capture_output=True, text=True)

    # Parse errors
    errors = parse_tsc_errors(proc.stdout)
    error_summary = format_error_summary(errors)

    # Determine success/failure
    success = proc.returncode == 0
    status = "PASS" if success else "FAIL"

    return {
        "success": success,
        "status": status,
        "returncode": proc.returncode,
        "errors": errors,
        "error_count": len(errors),
        "error_summary": error_summary,
        "stdout": proc.stdout,
        "stderr": proc.stderr,
        "temp_dir": td,
        "ts_file": ts_path,
        "tsconfig": os.path.join(td, "tsconfig.json"),
        "cmd": " ".join(cmd),
        "use_preview": use_preview
    }


# Example usage
if __name__ == "__main__":
    import sys
    import io
    # Fix Windows console encoding
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

    print("="*70)
    print("Optimized TypeScript Checker - Using Persistent Environment")
    print("="*70)

    # Test snippets
    snippets = [
        """
        /// <reference types="office-js" />
        async function test1() {
          await Word.run(async (context) => {
            const p = context.document.body.insertParagraph("Hello", Word.InsertLocation.start);
            await context.sync();
          });
        }
        """,
        """
        /// <reference types="office-js" />
        async function test2() {
          await Word.run(async (context) => {
            const p = context.document.body.insertParagraph("Test", Word.InsertLocation.start);
            p.font.bold = true;
            await context.sync();
          });
        }
        """,
        """
        /// <reference types="office-js" />
        async function testInvalid() {
          await Word.run(async (context) => {
            const p = context.document.body.insertParagraph("Test", Word.InsertLocation.start);
            p.nonExistentProperty = "fail";
            await context.sync();
          });
        }
        """
    ]

    import time
    start = time.time()

    for i, snippet in enumerate(snippets, 1):
        print(f"\n[Test {i}] Checking snippet...")
        result = check_officejs_ts_fast(snippet, use_preview=True)
        print(f"{'✅ PASS' if result['success'] else '❌ FAIL'} - {result['error_count']} error(s)")
        if not result['success']:
            print(f"{result['error_summary']}")

    elapsed = time.time() - start
    print(f"\n{'='*70}")
    print(f"Total time: {elapsed:.2f}s (avg {elapsed/len(snippets):.2f}s per check)")
    print(f"{'='*70}")

    # Cleanup
    cleanup_persistent_env()
