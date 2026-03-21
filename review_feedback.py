#!/usr/bin/env python3
"""
Feedback Review Tool for Claude Code
读取反馈数据，分析哪些值得处理，输出可操作的改进建议。
用法: python review_feedback.py [--url URL]
"""
import json, sys, os

FEEDBACK_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'feedback.json')

# 可操作的分类 - 只有这些才会被标记为需要处理
ACTIONABLE_CATS = {'data', 'format', 'feature'}

# 危险关键词 - 含这些内容的反馈自动忽略
DANGER_KEYWORDS = [
    'drop table', 'delete all', 'rm -rf', 'shutdown', 'exec(', 'eval(',
    'import os', '__import__', 'system(', 'subprocess', 'sql injection',
    '<script', 'javascript:', 'onerror', 'onclick',
]

def load_feedback(path=None):
    path = path or FEEDBACK_FILE
    if not os.path.exists(path):
        return []
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def is_dangerous(msg):
    lower = msg.lower()
    return any(kw in lower for kw in DANGER_KEYWORDS)

def is_actionable(fb):
    """判断反馈是否可操作"""
    msg = fb.get('message', '')
    cat = fb.get('category', '')
    # 已处理的跳过
    if fb.get('status') == 'resolved':
        return False
    # 危险内容
    if is_dangerous(msg):
        return False
    # 太短没有信息量
    if len(msg.strip()) < 5:
        return False
    # 可操作分类
    if cat in ACTIONABLE_CATS:
        return True
    # "其他"分类中包含数据/精度相关关键词也可操作
    data_keywords = ['精度', '准确', '错误', '不对', '算错', '格式', '缺少',
                     '多了', '少了', '重复', '遗漏', '公式', '计算', '数值',
                     'CBM', '箱', 'box', 'size', '尺码', '颜色', 'color',
                     'PI', 'PO', 'CI', 'COG', 'packing']
    lower = msg.lower()
    return any(kw.lower() in lower for kw in data_keywords)

def review():
    feedbacks = load_feedback()
    if not feedbacks:
        print("无反馈数据")
        return

    pending = [fb for fb in feedbacks if fb.get('status') == 'pending']
    actionable = [fb for fb in pending if is_actionable(fb)]
    ignored = [fb for fb in pending if not is_actionable(fb)]

    print(f"=== EMILY 反馈审核 ===")
    print(f"总数: {len(feedbacks)} | 待处理: {len(pending)} | 可操作: {len(actionable)} | 忽略: {len(ignored)}")
    print()

    if actionable:
        print("--- 可操作反馈 ---")
        for fb in actionable:
            print(f"  [{fb['id']}] [{fb['category']}] {fb['time']}")
            print(f"      {fb['message'][:200]}")
            print()

    if ignored:
        print(f"--- 已忽略 {len(ignored)} 条 ---")
        for fb in ignored:
            reason = "危险内容" if is_dangerous(fb['message']) else "不可操作/无关"
            print(f"  [{fb['id']}] {reason}: {fb['message'][:80]}")

def mark_resolved(fb_id):
    feedbacks = load_feedback()
    for fb in feedbacks:
        if fb['id'] == fb_id:
            fb['status'] = 'resolved'
    with open(FEEDBACK_FILE, 'w', encoding='utf-8') as f:
        json.dump(feedbacks, f, ensure_ascii=False, indent=2)
    print(f"反馈 #{fb_id} 已标记为已处理")

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--resolve':
        mark_resolved(int(sys.argv[2]))
    else:
        review()
