# -*- coding: utf-8 -*-
"""Weekly orchestrator: refresh email index -> fetch index closes (best effort)
-> write completed-week NAV rows into the master (with backup). Logs everything
to logs/run_YYYYMMDD_HHMMSS.log.

Usage:
    python run_weekly.py             # safe: preview only (不动正式表, 邮件只打印)
    python run_weekly.py --commit    # writes the master (with backup) + 发周报邮件
    python run_weekly.py --commit --no-notify   # 只写数据, 不发邮件(周中那几次用)
"""
import os, sys, subprocess, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable

def main():
    commit = "--commit" in sys.argv
    no_notify = "--no-notify" in sys.argv   # write silently (周中跑, 不打扰上司)
    os.makedirs(os.path.join(HERE, "logs"), exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    logf = os.path.join(HERE, "logs", f"run_{ts}.log")
    steps = [["build_index.py"], ["fill_index.py"]]
    if datetime.date.today().isocalendar()[1] % 2 == 0:   # 双周回归校验
        steps.append(["validate.py"])
    # 新链路: write/obs 只把新行算进一次性预览副本, 再由 com_sync 用 COM 移植进正式表
    # (不再用 openpyxl 直接写正式表 -> 不毁手工格式、不复活主题色)。apply.py 串起这三步。
    steps.append(["apply.py", "--weekly", "--obs"] + (["--commit"] if commit else []))
    if not no_notify:
        steps.append(["notify.py"] + ([] if commit else ["--dry"]))  # 预览模式只打印不发信
    NONFATAL = {"fill_index.py", "validate.py", "notify.py", "apply.py"}
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    with open(logf, "w", encoding="utf-8") as log:
        for step in steps:
            head = f"\n{'='*60}\n# {datetime.datetime.now():%H:%M:%S}  {' '.join(step)}\n{'='*60}\n"
            print(head.strip()); log.write(head)
            try:
                out = subprocess.run([PY] + [os.path.join(HERE, step[0])] + step[1:],
                                     cwd=HERE, env=env, capture_output=True, text=True,
                                     encoding="utf-8", errors="replace", timeout=600)
                print(out.stdout)
                log.write(out.stdout)
                if out.returncode != 0:
                    log.write("\n[STDERR]\n" + out.stderr)
                    print("[STDERR]\n" + out.stderr)
                    if step[0] not in NONFATAL:   # build_index / write are fatal
                        print(f"步骤 {step[0]} 失败，终止。详见 {logf}"); break
            except Exception as e:
                log.write(f"\n[EXCEPTION] {e}\n"); print(f"[EXCEPTION] {e}")
                if step[0] not in NONFATAL:
                    break
    print(f"\n日志: {logf}")

if __name__ == "__main__":
    main()
