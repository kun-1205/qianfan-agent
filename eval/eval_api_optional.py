"""智能体效果评估 —— API 自动化版（进阶可选）

============================================================
⚠️ 用之前先想清楚：你真的需要它吗？

  如果你只有 65 条用例，用 Excel 手测（见 eval_script.py）更快。
  本脚本适合：用例上百条、需要反复跑、或想接入自动化流程。

【它做什么 / 不做什么】

  ✅ 做：批量把问题发给智能体，记录「回答内容」和「响应耗时」
  ❌ 不做：自动评分

  为什么不自动评分？因为"答得对不对"目前只有人能可靠判断。
  脚本跑完会生成一张表，你仍然要人工填 score 列。
  这样做出来的报告才经得起面试官追问。

【准备工作（三样）】

  1. Python 3.8+
  2. 安装官方 SDK：
         pip install appbuilder-sdk
  3. 拿到两个值（千帆 AppBuilder 控制台 → 应用 → 发布 → API 信息）：
         APPBUILDER_API_KEY   你的 API Key
         智能体应用 ID         形如 xxxxxxxx-xxxx-xxxx

     把它们设为环境变量（不要把密钥写进代码！）：
         Windows PowerShell:
             $env:APPBUILDER_API_KEY="你的key"
         Windows CMD:
             set APPBUILDER_API_KEY=你的key
         macOS / Linux:
             export APPBUILDER_API_KEY="你的key"

【运行】

  # 先测一条，确认能连通
  python eval_api_optional.py --test

  # 确认没问题后，跑全部用例
  python eval_api_optional.py

  结果会写入 eval/api_result.csv（含回答内容，score 列留空待填）

【重要提醒】
  · 本脚本调用的 SDK 接口可能随官方版本变化，
    若报错请以官方文档为准：https://cloud.baidu.com/doc/qianfan/s/Gmh4suel1
  · 绝对不要把 API Key 提交到 GitHub
============================================================
"""

import argparse
import csv
import os
import sys
import time

# ---------- 路径 ----------
EVAL_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(EVAL_DIR)
CASES_FILE = os.path.join(BASE_DIR, "tests", "test_cases.csv")
RESULT_FILE = os.path.join(EVAL_DIR, "api_result.csv")

# ---------- 从环境变量读取配置（不要在代码里写密钥） ----------
API_KEY = os.getenv("APPBUILDER_API_KEY")
APP_ID = os.getenv("APPBUILDER_APP_ID")


def check_env() -> bool:
    """检查环境和配置是否就绪，缺什么就明确告诉用户缺什么。"""
    ok = True

    # 1. SDK 是否安装
    try:
        import appbuilder  # noqa: F401
    except ImportError:
        print("❌ 未安装 SDK。请先运行：")
        print("     pip install appbuilder-sdk")
        print()
        print("   如果安装失败（网络问题），可以试试国内镜像：")
        print("     pip install appbuilder-sdk -i https://pypi.tuna.tsinghua.edu.cn/simple")
        ok = False

    # 2. API Key
    if not API_KEY:
        print("❌ 未找到环境变量 APPBUILDER_API_KEY")
        print("   请按脚本开头的说明设置环境变量。")
        print("   提示：不要把它直接写进代码，否则可能被传到 GitHub 上泄露。")
        ok = False

    # 3. 应用 ID
    if not APP_ID:
        print("⚠️  未设置 APPBUILDER_APP_ID（部分接口需要）")
        print("   如果你的调用报错说缺 app_id，请补上这个环境变量。")

    return ok


def get_client():
    """初始化智能体客户端。

    ⚠️ 不同 SDK 版本的类名/调用方式可能不同，若此处报错，
       请对照官方文档调整：
       https://cloud.baidu.com/doc/qianfan/s/Gmh4suel1
    """
    from appbuilder import AppBuilderClient

    return AppBuilderClient(api_key=API_KEY, app_id=APP_ID)


def ask_once(client, question: str, conversation_id: str = None):
    """向智能体提一个问题，返回 (回答文本, 会话ID)。

    stream=False 表示一次性返回全部内容，便于本脚本记录。
    """
    if conversation_id is None:
        conversation_id = client.create_conversation()

    message = client.run(conversation_id, query=question, stream=False)

    # 不同版本返回结构可能不同，这里做兼容处理：
    # 优先取常见的 answer / content 字段，取不到就把对象转成字符串
    answer = ""
    for attr in ("answer", "content", "text"):
        if hasattr(message, attr):
            answer = str(getattr(message, attr))
            break
    if not answer:
        # 有些版本把内容放在 content 属性里的列表结构中
        try:
            answer = message.content.answer
        except Exception:
            answer = str(message)

    return answer, conversation_id


def run_test() -> None:
    """只跑第一条用例，用来验证配置是否正确。"""
    print("=" * 52)
    print("  连通性测试（只发 1 条）")
    print("=" * 52)

    if not check_env():
        sys.exit(1)

    client = get_client()
    question = "你好，请简单介绍一下你自己。"  # 用一个简单问题测试连通

    print(f"\n正在发送测试问题：{question}")
    start = time.time()
    try:
        answer, _ = ask_once(client, question)
        cost = round(time.time() - start, 2)
        print(f"\n✅ 调用成功（耗时 {cost} 秒）")
        print("-" * 52)
        print(answer[:500])
        print("-" * 52)
        print("\n配置没问题，可以跑全部用例了：")
        print("    python eval_api_optional.py")
    except Exception as e:
        print(f"\n❌ 调用失败：{type(e).__name__}: {e}")
        print()
        print("排查建议：")
        print("  1. 检查 API Key 是否正确、是否已过期")
        print("  2. 检查该智能体是否已在千帆控制台「发布」")
        print("  3. 检查账号是否已开通 AppBuilder 服务")
        print("  4. 若提示参数错误，请对照官方文档确认 SDK 用法")
        sys.exit(1)


def run_all() -> None:
    """批量跑全部用例，记录回答与耗时。"""
    print("=" * 52)
    print("  批量评测（自动提问，人工评分）")
    print("=" * 52)

    if not check_env():
        sys.exit(1)

    if not os.path.exists(CASES_FILE):
        print(f"❌ 找不到用例文件：{CASES_FILE}")
        sys.exit(1)

    with open(CASES_FILE, encoding="utf-8-sig", newline="") as f:
        cases = list(csv.DictReader(f))

    print(f"\n共 {len(cases)} 条用例。")
    print("提示：这会消耗 API 额度，且每条可能耗时数秒到数十秒。\n")

    client = get_client()
    rows = []

    for i, case in enumerate(cases, 1):
        q = case["question"]
        print(f"[{i}/{len(cases)}] {q[:36]}...", end=" ", flush=True)

        start = time.time()
        try:
            answer, _ = ask_once(client, q)
            cost = round(time.time() - start, 2)
            print(f"✅ {cost}s")
        except Exception as e:
            answer = f"[调用失败] {type(e).__name__}: {e}"
            cost = round(time.time() - start, 2)
            print(f"❌ {cost}s")

        rows.append({
            "id": case["id"],
            "category": case["category"],
            "question": q,
            "expected_point": case.get("expected_point", ""),
            "is_answerable": case.get("is_answerable", ""),
            "answer": answer,
            "latency_s": cost,
            "score": "",     # ← 留空，等你自己判断填 0/1/2
            "note": "",
        })

        time.sleep(0.5)  # 稍微歇一下，避免请求过密

    # 写出结果
    with open(RESULT_FILE, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    print()
    print(f"✅ 结果已写入：{RESULT_FILE}")
    print()
    print("下一步：")
    print("  1. 用 Excel 打开这个文件，对照 answer 列检查每个回答")
    print("  2. 在 score 列填 0（错）/ 1（部分对）/ 2（完全对）")
    print("  3. 想生成统计报告的话，把这个文件另存为")
    print("     eval/human_eval.csv 并删掉 answer 列，再运行 eval_script.py")


def main() -> None:
    parser = argparse.ArgumentParser(description="智能体 API 批量评测")
    parser.add_argument("--test", action="store_true", help="只测一条，验证配置")
    args = parser.parse_args()

    if args.test:
        run_test()
    else:
        run_all()


if __name__ == "__main__":
    main()
