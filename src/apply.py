# -*- coding: utf-8 -*-
"""新写入链路的总成: 把"算新行"和"COM 移植进正式表"串起来。

  ① 复制正式表 -> 一次性预览副本
  ② write.py --book 预览   (周更产品的新行算进预览; 复杂逻辑原样)
  ③ obs_daily.py --book 预览 (观察仓的新行算进预览)
  ④ com_sync.py --preview 预览 [--commit]  (用 COM 把预览里多出的新行/补的空格移植进正式表)

好处: write/obs 不再用 openpyxl 碰正式表 -> 不毁手工格式、不复活主题色(到 Linda 电脑变色);
真正写正式表的只有 com_sync 一处(COM), 格式保留、写死 RGB 红绿。

用法: apply.py [--weekly] [--obs] [--commit] [--preview PATH]
  不指定 --weekly/--obs 时默认两者都做; 不加 --commit = 干跑(生成预览+报告, 不动正式表)。
"""
import os, sys, shutil, subprocess
import navlib as L

HERE = L.HERE
PY = sys.executable

def run(args):
    print(f"\n----- {' '.join(args)} -----", flush=True)
    r = subprocess.run([PY, os.path.join(HERE, args[0])] + args[1:], cwd=HERE,
                       env=dict(os.environ, PYTHONIOENCODING="utf-8"),
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=600)
    print(r.stdout)
    if r.returncode != 0:
        print("[STDERR]\n" + (r.stderr or ""))
    return r.returncode

def main():
    commit = "--commit" in sys.argv
    weekly = "--weekly" in sys.argv
    obs = "--obs" in sys.argv
    if not weekly and not obs:
        weekly = obs = True
    master = L.CFG["master_path"]
    if "--preview" in sys.argv:
        preview = sys.argv[sys.argv.index("--preview") + 1]
    else:
        d = os.path.dirname(master)
        stem = os.path.splitext(os.path.basename(master))[0]
        preview = os.path.join(d, stem + "_自动更新预览.xlsx")

    shutil.copy2(master, preview)               # 预览副本 = 正式表的当前拷贝
    print(f"预览副本(正式表拷贝): {preview}")
    if weekly:
        run(["write.py", "--book", preview])    # 单步失败不致命: 预览里没算到的, com_sync 自然不会搬
    if obs and os.path.exists(os.path.join(HERE, "obs_daily.py")):  # 无观察仓模块(如纯周更部署)则跳过
        run(["obs_daily.py", "--book", preview])
    return run(["com_sync.py", "--preview", preview] + (["--commit"] if commit else []))

if __name__ == "__main__":
    sys.exit(main())
