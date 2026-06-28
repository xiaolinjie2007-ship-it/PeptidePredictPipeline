#!/usr/bin/env python3
"""
PeptideRanker 完整流程（最优版，融合 V1/V2 全部改进）
=====================================================
一键运行：python full_pipeline.py
分批运行：python full_pipeline.py --batch 5

功能：
  1. 读取 digestion_results_normalized.xlsx（自动过滤长度<2的肽）
  2. 按 (protein_name, scheme_id) 分组
  3. 逐组提交到 PeptideRanker（Playwright → Node.js）
  4. 监控 QQ 邮箱（INBOX + 垃圾邮件），自动重连
  5. 下载 TSV 附件，解析 bioactivity_score
  6. 更新输出 Excel（仅标准库，无需 pandas）
  7. 断点续传 + 已处理邮件去重，防止张冠李戴
"""

import imaplib, email as em, os, json, time, shutil, subprocess, logging, zipfile
import xml.etree.ElementTree as ET
from email.header import decode_header
from datetime import datetime
from config import EMAIL_ADDR, EMAIL_PASS, IMAP_SERVER, IMAP_PORT

# ============================================================
# 配置
# ============================================================
CONFIG = {
    "email_addr":   EMAIL_ADDR,
    "email_pass":   EMAIL_PASS,
    "imap_server":  IMAP_SERVER,
    "imap_port":    IMAP_PORT,
    "excel_input":  r"c:\Users\kanmao\.claude\skills\digestion_results_normalized.xlsx",
    "excel_output": r"c:\Users\kanmao\Desktop\digestion_results_with_bioactivity.xlsx",
    "node_script":  r"c:\Users\kanmao\.claude\skills\submit_peptideranker.js",
    "download_dir": r"c:\Users\kanmao\Desktop\PeptideRanker_Results",
    "state_file":   r"c:\Users\kanmao\.claude\skills\pipeline_state.json",
    "log_file":     r"c:\Users\kanmao\.claude\skills\pipeline.log",
    "poll_interval": 45,
    "max_wait":      120,
}
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

# ============================================================
# 日志
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(), logging.FileHandler(CONFIG["log_file"], encoding="utf-8")],
)
log = logging.getLogger("Pipeline")

# ============================================================
# Excel 操作（纯标准库 XML 解析）
# ============================================================
def _cell_text(c):
    """读取单元格文本（内联字符串）"""
    is_el = c.find(f'{NS}is')
    if is_el is not None:
        t = is_el.find(f'{NS}t')
        return t.text if t is not None and t.text else ''
    v = c.find(f'{NS}v')
    return v.text if v is not None else ''

def _col_letter(ref):
    """从单元格引用（如 'A1'）提取列字母"""
    return ''.join(filter(str.isalpha, ref))

def read_excel_groups(path):
    """读取 Excel，筛选长度≥2 的肽段，按 (protein, scheme_id) 分组"""
    log.info("读取 Excel: %s", path)
    with zipfile.ZipFile(path, 'r') as zf:
        sheet = ET.parse(zf.open('xl/worksheets/sheet1.xml'))
        rows = sheet.findall(f'.//{NS}row')
        groups = {}
        total = kept = 0
        for r in rows[1:]:
            cells = {}
            for c in r.findall(f'{NS}c'):
                cells[_col_letter(c.attrib.get('r', ''))] = _cell_text(c)
            total += 1
            try:
                if int(cells.get('K', '0')) >= 2:
                    key = (cells.get('B', ''), cells.get('E', ''))
                    groups.setdefault(key, []).append(cells.get('J', ''))
                    kept += 1
            except (ValueError, TypeError):
                pass
    log.info("  共 %d 行，%d 条肽段，%d 组", total, kept, len(groups))
    return groups

def update_excel(input_path, output_path, predictions):
    """将 bioactivity_score 写入 Excel 新列 L"""
    with zipfile.ZipFile(input_path, 'r') as zf_in:
        sheet_xml = zf_in.read('xl/worksheets/sheet1.xml').decode('utf-8')
        root = ET.fromstring(sheet_xml)
        all_rows = root.findall(f'.//{NS}row')
        if not all_rows:
            return

        new_col = 'L'
        # 表头
        hdr = all_rows[0]
        nc = ET.SubElement(hdr, f'{NS}c', {'r': f'{new_col}1', 't': 'inlineStr'})
        is_el = ET.SubElement(nc, f'{NS}is')
        t_el = ET.SubElement(is_el, f'{NS}t')
        t_el.text = 'bioactivity_score'
        t_el.set(XML_SPACE, 'preserve')

        # 构建 行号→分数 映射（按 J 列肽序列匹配）
        row_score = {}
        for r in all_rows[1:]:
            rn = int(r.attrib.get('r', '0'))
            for c in r.findall(f'{NS}c'):
                if _col_letter(c.attrib.get('r', '')) == 'J':
                    is_el = c.find(f'{NS}is')
                    if is_el is not None:
                        t = is_el.find(f'{NS}t')
                        seq = t.text if t is not None and t.text else ''
                        if seq in predictions:
                            row_score[rn] = predictions[seq]
                    break

        # 写入
        for r in all_rows[1:]:
            rn = int(r.attrib.get('r', '0'))
            nc = ET.SubElement(r, f'{NS}c', {'r': f'{new_col}{rn}', 't': 'inlineStr'})
            is_el = ET.SubElement(nc, f'{NS}is')
            t_el = ET.SubElement(is_el, f'{NS}t')
            t_el.text = str(row_score.get(rn, ''))
            t_el.set(XML_SPACE, 'preserve')

        modified = ET.tostring(root, encoding='unicode')
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.infolist():
                zf_out.writestr(item,
                    modified if item.filename == 'xl/worksheets/sheet1.xml'
                    else zf_in.read(item.filename))

# ============================================================
# IMAP 邮件操作
# ============================================================
def imap_connect():
    """连接 IMAP，最多重试 5 次"""
    for i in range(5):
        try:
            mail = imaplib.IMAP4_SSL(CONFIG["imap_server"], CONFIG["imap_port"], timeout=30)
            mail.login(CONFIG["email_addr"], CONFIG["email_pass"])
            return mail
        except Exception as e:
            log.warning("IMAP 连接失败 (%d/5): %s", i+1, e)
            time.sleep(10)
    raise ConnectionError("IMAP 连接失败（5 次重试）")

def find_result_emails(mail, since_uid):
    """
    搜索 INBOX + 垃圾邮件中 PeptideRanker 结果邮件
    返回 [(uid_bytes, folder_name), ...]
    如果两个文件夹全部失败则抛出 ConnectionError 触发重连
    """
    results = []
    failures = 0
    for folder in ["INBOX", "Junk"]:
        try:
            mail.select(folder)
            status, data = mail.uid("SEARCH", None, "ALL")
            if status != "OK" or not data or not data[0]:
                continue
            all_uids = data[0].split()
            new = [u for u in all_uids if int(u) > since_uid]
            for uid in new[-50:]:
                try:
                    status, md = mail.uid("FETCH", uid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT)])")
                    if status != "OK":
                        continue
                    for item in md:
                        if isinstance(item, tuple) and len(item) > 1:
                            if "PeptideRanker Results:" in item[1].decode("utf-8", errors="replace"):
                                results.append((uid, folder))
                except Exception:
                    pass
        except (imaplib.IMAP4.abort, imaplib.IMAP4.error, ConnectionError, OSError) as e:
            log.warning("搜索文件夹 %s 出错: %s", folder, e)
            failures += 1
            continue
    # V2 关键改进：两个文件夹都失败 → 通知调用方重连
    if failures >= 2:
        raise ConnectionError("IMAP 连接断开（所有文件夹均失败）")
    return results

def download_attachment(mail, uid, folder="INBOX"):
    """
    下载邮件中的 TSV 附件，返回文件路径或 None
    IMAP 异常会向上抛出让调用方重连（V2 改进）
    """
    mail.select(folder)
    status, data = mail.uid("FETCH", uid, "(RFC822)")
    if status != "OK" or not data:
        return None
    raw = None
    for item in data:
        if isinstance(item, tuple) and len(item) > 1:
            raw = item[1]; break
    if raw is None:
        return None

    msg = em.message_from_bytes(raw)
    for part in msg.walk():
        if "attachment" in str(part.get("Content-Disposition", "")):
            fn = part.get_filename()
            if fn:
                parts = decode_header(fn)
                if parts:
                    fn = parts[0][0]
                    if isinstance(fn, bytes):
                        fn = fn.decode("utf-8", errors="replace")
                fp = os.path.join(CONFIG["download_dir"], fn)
                with open(fp, "wb") as f:
                    f.write(part.get_payload(decode=True))
                return fp
    return None

def parse_tsv(filepath):
    """解析 PeptideRanker TSV → {sequence: prediction}"""
    preds = {}
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        f.readline()
        for line in f:
            line = line.strip()
            if not line: continue
            parts = line.split('\t')
            if len(parts) >= 2:
                preds[parts[0].strip()] = parts[1].strip()
    return preds

# ============================================================
# PeptideRanker 提交（Playwright → Node.js）
# ============================================================
def submit_peptides(peptides_str, email_addr):
    """提交肽段，返回 (success, error)，最多重试 3 次"""
    escaped = peptides_str.replace('\\', '\\\\').replace('"', '\\"')
    cmd = ["node", CONFIG["node_script"], escaped, email_addr]
    for i in range(3):
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=120,
                               cwd=os.path.dirname(CONFIG["node_script"]))
            out = r.stdout.strip()
            if out:
                d = json.loads(out)
                if d.get("success"): return True, None
                return False, d.get("error", "未知错误")
            return False, f"无输出，stderr: {r.stderr[:200]}"
        except subprocess.TimeoutExpired:
            log.warning("提交超时 (%d/3)", i+1)
        except Exception as e:
            log.warning("提交异常 (%d/3): %s", i+1, e)
        time.sleep(5)
    return False, "3 次尝试均失败"

# ============================================================
# 状态持久化
# ============================================================
def load_state():
    if os.path.exists(CONFIG["state_file"]):
        with open(CONFIG["state_file"], 'r') as f:
            return json.load(f)
    return {
        "processed_group_index": 0,
        "all_predictions": {},
        "baseline_uid": 0,
        "processed_emails": [],    # V2 关键：防重复处理
    }

def save_state(state):
    with open(CONFIG["state_file"], 'w') as f:
        json.dump(state, f, indent=2, ensure_ascii=False)

# ============================================================
# 主流程
# ============================================================
def main(batch_size=None):
    log.info("=" * 60)
    log.info("PeptideRanker 完整流程  @ %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    log.info("=" * 60)

    # ---- 1. 读取数据 ----
    groups = read_excel_groups(CONFIG["excel_input"])
    group_keys = sorted(groups.keys(), key=lambda x: (x[0], int(x[1])))
    log.info("共 %d 组，%d 条肽段", len(group_keys), sum(len(v) for v in groups.values()))

    # ---- 2. 恢复状态 ----
    state = load_state()
    start_idx    = state["processed_group_index"]
    all_preds    = state["all_predictions"]
    baseline_uid = state.get("baseline_uid", 0)
    processed    = set(state.get("processed_emails", []))  # V2：已处理的邮件 UID

    # ---- 3. 准备输出 ----
    if not os.path.exists(CONFIG["excel_output"]):
        shutil.copy2(CONFIG["excel_input"], CONFIG["excel_output"])
    os.makedirs(CONFIG["download_dir"], exist_ok=True)

    # ---- 4. 初始化 baseline UID ----
    if baseline_uid == 0:
        mail = imap_connect()
        max_uid = 0
        for folder in ["INBOX", "Junk"]:
            try:
                mail.select(folder)
                status, data = mail.uid("SEARCH", None, "ALL")
                if status == "OK" and data and data[0]:
                    max_uid = max(max_uid, int(data[0].split()[-1]))
            except Exception:
                pass
        baseline_uid = max_uid
        state["baseline_uid"] = baseline_uid
        save_state(state)
        mail.logout()
    log.info("起点: 第 %d/%d 组, baseline UID=%d", start_idx+1, len(group_keys), baseline_uid)

    if start_idx >= len(group_keys):
        log.info("全部完成！")
        return

    end_idx = len(group_keys) if batch_size is None else min(start_idx + batch_size, len(group_keys))
    log.info("本轮: %d ~ %d 组（%d 组）", start_idx+1, end_idx, end_idx - start_idx)

    # ---- 5. 连接 IMAP ----
    mail = imap_connect()

    # ---- 6. 逐组处理 ----
    for idx in range(start_idx, end_idx):
        protein, scheme_id = group_keys[idx]
        peptides = groups[(protein, scheme_id)]
        label = f"{protein} / 方案 {scheme_id}"
        log.info("[%d/%d] %s (%d 条肽段)", idx+1, len(group_keys), label, len(peptides))

        # -- 6a. 提交 --
        ok, err = submit_peptides("\n".join(peptides), CONFIG["email_addr"])
        if not ok:
            log.error("  ✗ 提交失败: %s", err)
            save_state(state)
            continue
        log.info("  ✓ 已提交，等待邮件...")

        # -- 6b. 等待结果 --
        t0 = time.time()
        result_file = None
        while time.time() - t0 < CONFIG["max_wait"] * 60:
            time.sleep(CONFIG["poll_interval"])

            # 搜索邮件（V2：全部文件夹失败 → 抛异常 → 重连）
            try:
                new_emails = find_result_emails(mail, baseline_uid)
            except Exception:
                try: mail.logout()
                except Exception: pass
                mail = imap_connect()
                continue

            for euid, folder in new_emails:
                uid_str = euid.decode()
                # V2 关键：跳过已处理的邮件，防止延迟邮件被错误匹配
                if uid_str in processed:
                    continue
                processed.add(uid_str)
                # 下载（V2：异常 → 重连后重试一次）
                try:
                    result_file = download_attachment(mail, euid, folder)
                except Exception:
                    mail = imap_connect()
                    try:
                        result_file = download_attachment(mail, euid, folder)
                    except Exception:
                        result_file = None
                if result_file:
                    log.info("  ✓ 收到: %s", os.path.basename(result_file))
                    break

            if result_file:
                break

            elapsed = int(time.time() - t0)
            if elapsed % 300 < CONFIG["poll_interval"]:
                log.info("  等待中... %d 分钟", elapsed // 60)

        if not result_file:
            log.error("  ✗ 超时 (%d 分钟)", CONFIG["max_wait"])
        else:
            # -- 6c. 解析 & 合并 --
            preds = parse_tsv(result_file)
            log.info("  解析 %d 条预测", len(preds))
            all_preds.update(preds)
            # -- 6d. 更新 Excel --
            try:
                update_excel(CONFIG["excel_input"], CONFIG["excel_output"], all_preds)
                log.info("  ✓ Excel 已更新（累计 %d 条）", len(all_preds))
            except Exception as e:
                log.error("  ✗ Excel 更新失败: %s", e)

        # -- 6e. 保存进度 --
        state["processed_group_index"] = idx + 1
        state["all_predictions"] = all_preds
        state["processed_emails"] = list(processed)
        save_state(state)

    # ---- 7. 收尾 ----
    try: mail.logout()
    except Exception: pass
    log.info("=" * 60)
    log.info("本轮完成！%d/%d (%d%%)　累计预测: %d",
             state["processed_group_index"], len(group_keys),
             state["processed_group_index"]*100//len(group_keys), len(all_preds))
    log.info("输出: %s", CONFIG["excel_output"])

if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="PeptideRanker 完整流程")
    p.add_argument("--batch", type=int, help="只处理 N 组后退出")
    args = p.parse_args()
    try:
        main(batch_size=args.batch)
    except KeyboardInterrupt:
        log.info("用户中断，状态已保存")
    except Exception as e:
        log.error("致命错误: %s", e, exc_info=True)
