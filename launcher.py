# launcher.py - the small entry point of the .exe. It rarely changes.
# It starts the newest installed update (folder %APPDATA%\MobilePOS_V5\app\<version>) or, if there is none, the version built into the .exe.
# If a new update does not start properly twice, it goes back to the previous version by itself.
import os, sys, json, time, importlib.util, traceback

BUILTIN_VERSION = "V5.20"          # the version this .exe was built with
APP_NAME = "MobilePOS_V5"
DATA_DIR = os.environ.get("MOBILEPOS_DATA") or os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), APP_NAME)
APP_DIR = os.path.join(DATA_DIR, "app")
PTR = os.path.join(APP_DIR, "current.json")
FILES = ("pos_core.py", "pos_server.py", "index.html")
DEFAULT = {"version": "builtin", "prev": None, "status": "ok", "tries": 0}


def log(msg):
    try:
        os.makedirs(APP_DIR, exist_ok=True)
        with open(os.path.join(APP_DIR, "launcher.log"), "a", encoding="utf-8") as f: f.write(time.strftime("%Y-%m-%d %H:%M:%S ") + msg + "\n")
    except Exception: pass


def read_ptr():
    try:
        with open(PTR, encoding="utf-8") as f: return {**DEFAULT, **json.load(f)}
    except Exception: return dict(DEFAULT)


def write_ptr(p):
    try:
        os.makedirs(APP_DIR, exist_ok=True)
        with open(PTR + ".tmp", "w", encoding="utf-8") as f: json.dump(p, f)
        os.replace(PTR + ".tmp", PTR)
    except Exception: log("could not write current.json: " + traceback.format_exc(limit=1))


def complete(folder): return all(os.path.isfile(os.path.join(folder, f)) for f in FILES)


def choose(depth=0):
    p = read_ptr(); v = p.get("version", "builtin")
    if v == "builtin" or depth > 3: return None
    folder = os.path.join(APP_DIR, v)
    if not complete(folder):
        log(f"{v} is missing files - using the built-in version"); write_ptr({**DEFAULT, "note": f"{v} was incomplete; the built-in version is used"}); return None
    if p.get("status") == "new":
        p["tries"] = int(p.get("tries", 0)) + 1
        if p["tries"] > 2:                      # the new version never confirmed that it started well: go back
            back = p.get("prev") or "builtin"
            log(f"{v} did not start properly twice - going back to {back}")
            write_ptr({"version": back, "prev": None, "status": "ok", "tries": 0, "note": f"{v} did not start properly and was rolled back to {back if back != 'builtin' else 'the built-in version'}"})
            return choose(depth + 1)
        write_ptr(p)
    return folder


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path); mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod; spec.loader.exec_module(mod); return mod


def fall_back_to_builtin(reason):
    log("falling back to the built-in version: " + reason)
    p = read_ptr(); write_ptr({**DEFAULT, "note": f"{p.get('version')} could not start and the built-in version is used instead"})
    for n in ("pos_core", "pos_server"): sys.modules.pop(n, None)
    os.environ.pop("POS_APP_DIR", None)


def start():
    try: time.sleep(float(os.environ.pop("POS_RESTART_DELAY", "0") or 0))
    except Exception: pass
    os.environ["POS_BUILTIN_VERSION"] = BUILTIN_VERSION
    os.environ["POS_LAUNCHER"] = os.path.abspath(__file__)
    folder = choose(); srv = None
    if folder:
        try:
            os.environ["POS_APP_DIR"] = folder
            load("pos_core", os.path.join(folder, "pos_core.py")); srv = load("pos_server", os.path.join(folder, "pos_server.py"))
        except Exception:
            fall_back_to_builtin(traceback.format_exc(limit=3)); srv = None
    if srv is None:
        import pos_server as srv
    if os.environ.get("POS_LAUNCH_TEST"):
        print("LAUNCH", getattr(srv, "VERSION", "?"), os.environ.get("POS_APP_DIR", "builtin")); return
    srv.main()


if __name__ == "__main__":
    start()
