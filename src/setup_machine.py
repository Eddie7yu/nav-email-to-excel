# -*- coding: utf-8 -*-
"""一键安装：在新电脑（公司电脑）上配置好本工具。

用法（在 nav_tool 目录下）：
    python setup_machine.py --auth 邮箱授权码 [--master 净值表文件名或绝对路径] [--user 邮箱]

它会：
  1) 把邮箱授权码写到本机非同步位置 %LOCALAPPDATA%\\nav_tool\\secret.json
  2) （可选）更新 config.json 的净值表文件名 / 邮箱
  3) pip 安装依赖 (requirements.txt)
  4) 重新生成 registry.json（按当前净值表结构）
  5) 建立邮件索引 index.json
完成后用 `python run_weekly.py` 预览，确认无误再 `--commit`。
"""
import os, sys, json, argparse, subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable

def run(args):
    print("  $", " ".join(args))
    return subprocess.run(args, cwd=HERE, env=dict(os.environ, PYTHONIOENCODING="utf-8"))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--auth", required=True, help="QQ 邮箱 IMAP 授权码")
    ap.add_argument("--master", default=None, help="净值表文件名或绝对路径")
    ap.add_argument("--user", default=None, help="邮箱地址（默认沿用 config.json）")
    a = ap.parse_args()

    # 1) secret -> LOCALAPPDATA (非 OneDrive 同步)
    secdir = os.path.join(os.environ.get("LOCALAPPDATA", HERE), "nav_tool")
    os.makedirs(secdir, exist_ok=True)
    secp = os.path.join(secdir, "secret.json")
    json.dump({"password": a.auth}, open(secp, "w", encoding="utf-8"))
    print(f"[1/5] 授权码已写入 {secp}")

    # 2) config 更新
    cfgp = os.path.join(HERE, "config.json")
    cfg = json.load(open(cfgp, encoding="utf-8"))
    if a.master:
        cfg["master_path"] = a.master
    if a.user:
        cfg["imap"]["user"] = a.user
    cfg["imap"]["password"] = ""  # 不在 config 里存密码
    json.dump(cfg, open(cfgp, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"[2/5] 已更新 config.json (master_path={cfg['master_path']}, user={cfg['imap']['user']})")

    # 3) 依赖
    print("[3/5] 安装依赖 ...")
    run([PY, "-m", "pip", "install", "-r", os.path.join(HERE, "requirements.txt")])

    # 4) registry
    print("[4/5] 生成 registry.json ...")
    if run([PY, os.path.join(HERE, "build_registry.py")]).returncode != 0:
        print("  !! 生成 registry 失败，请检查 master_path 是否指向正确的净值表"); return

    # 5) index
    print("[5/5] 建立邮件索引 ...")
    if run([PY, os.path.join(HERE, "build_index.py")]).returncode != 0:
        print("  !! 建索引失败，请检查邮箱/授权码/网络"); return

    print("\n安装完成。建议先预览：  python run_weekly.py")
    print("确认无误后正式写入：     python run_weekly.py --commit")

if __name__ == "__main__":
    main()
