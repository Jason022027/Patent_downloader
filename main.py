import base64
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests
import pandas as pd
from tqdm import tqdm


# ====== 你要改的設定 ======
API_USER = "opdUser921"
API_PASS = "VHa5hsmN9x"

INPUT_FILE = "patents.xlsx"   # 也可用 patents.csv
INPUT_COLUMN = "公開公告號"

SAVE_DIR = Path("pdf_downloads")
SAVE_DIR.mkdir(exist_ok=True)

# 若你只想下載「公報」可以在這裡設關鍵字（保留 None 表示全抓）
# 例：只抓含「公報」的檔名：["公報"]
# 例：只抓 PDF：[".pdf"]
FILENAME_INCLUDE_KEYWORDS: Optional[List[str]] = None

# 速度與穩定性
SLEEP_BETWEEN_CASES = 0.25
RETRY = 3
TIMEOUT = 60


# ====== 官方 API endpoints ======
AUTH_URL = "https://tiponet.tipo.gov.tw/S092_API/opd1/getAuth"
CASEINFO_URL = "https://tiponet.tipo.gov.tw/S092_API/opd1/getCaseInfo/{}"
FILELIST_URL = "https://tiponet.tipo.gov.tw/S092_API/opd1/getResultFileList/{}"
DOWNLOAD_URL = "https://tiponet.tipo.gov.tw/S092_API/opd1/getfile/{}"


def _basic_auth_header(user: str, pwd: str) -> str:
    raw = f"{user}:{pwd}".encode("utf-8")
    return base64.b64encode(raw).decode("ascii")


def get_token(session: requests.Session) -> str:
    basic = _basic_auth_header(API_USER, API_PASS)
    r = session.get(
        AUTH_URL,
        headers={"Authorization": f"Basic {basic}"},
        timeout=TIMEOUT,
    )
    r.raise_for_status()

    # 正常情況會回傳 JSON，但有時候可能回傳非 JSON（例如純文字或 HTML），
    # 所以先嘗試解析 JSON，失敗時再用一些簡單的正則去抓 token。
    try:
        data = r.json()
    except requests.exceptions.JSONDecodeError:
        text = r.text or ""
        # 常見的 token 樣式：{"token":"..."} 或 {"access_token":"..."}，或 token=... 等
        m = re.search(r'"(?:access_)?token"\s*:\s*"([^"]+)"', text)
        if m:
            return m.group(1)
        m = re.search(r"token=([^&\s]+)", text)
        if m:
            return m.group(1)
        # 如果回傳是單純的字串（例如直接回傳 token），就回傳整個 trimmed 文本
        if text.strip():
            return text.strip()
        raise RuntimeError(f"getAuth 回傳無法解析的內容（非 JSON），前 300 字：{text[:300]!r}")

    token = None
    if isinstance(data, dict):
        token = data.get("token") or data.get("access_token") or data.get("data")
    elif isinstance(data, str) and data.strip():
        token = data.strip()

    if not token:
        raise RuntimeError(f"getAuth 回傳找不到 token：{data}")
    return token


def normalize_case_id(pubno: str) -> str:
    """
    常見格式：TW202528785A -> 202528785
    若不符合就原樣回傳（讓 API 自己判斷）
    """
    s = (pubno or "").strip()
    m = re.fullmatch(r"TW(\d+)[A-Z]\d*", s, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    return s


def safe_filename(name: str) -> str:
    # Windows 不允許的字元替換掉
    return re.sub(r'[<>:"/\\|?*\n\r\t]', "_", name).strip()


def keyword_filter(filename: str) -> bool:
    if FILENAME_INCLUDE_KEYWORDS is None:
        return True
    f = filename.lower()
    return any(k.lower() in f for k in FILENAME_INCLUDE_KEYWORDS)


def request_json_with_retry(
    session: requests.Session,
    method: str,
    url: str,
    headers: Dict[str, str],
) -> Dict[str, Any]:
    last_err = None
    for attempt in range(1, RETRY + 1):
        try:
            r = session.request(method, url, headers=headers, timeout=TIMEOUT)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            last_err = e
            time.sleep(0.6 * attempt)
    raise RuntimeError(f"請求失敗：{url}，最後錯誤：{last_err}")


def get_case_info(session: requests.Session, token: str, case_id: str) -> Dict[str, Any]:
    return request_json_with_retry(
        session,
        "GET",
        CASEINFO_URL.format(case_id),
        headers={"Authorization": f"Bearer {token}"},
    )


def get_file_list(session: requests.Session, token: str, case_no: str) -> Any:
    # 這支回傳可能是 list 或 dict（看系統版本），先用 Any 接住
    url = FILELIST_URL.format(case_no)
    last_err = None
    for attempt in range(1, RETRY + 1):
        try:
            r = session.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=TIMEOUT)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            last_err = e
            time.sleep(0.6 * attempt)
    raise RuntimeError(f"getResultFileList 失敗：{url}，最後錯誤：{last_err}")


def parse_file_items(file_list_json: Any) -> List[Tuple[str, str]]:
    """
    回傳 [(fileName, fileId或fileURL), ...]

    支援的格式包括：
    - list[dict]
    - dict 裡面包 list (data/files)
    - 特別處理 TIPO 新版回傳的 {"resultFileList": [...]} 內含 "fileList" 的結構
    """
    items: List[Tuple[str, str]] = []

    def _add(fname: Optional[str], fid: Optional[str]) -> None:
        if not fname or not fid:
            return
        items.append((str(fname), str(fid)))

    # 新版 API 回傳格式：{"resultFileList": [{..., "fileList":[{ "showName":..., "fileURL":...}, ...]}, ...]}
    if isinstance(file_list_json, dict) and file_list_json.get("resultFileList"):
        for entry in file_list_json.get("resultFileList", []) or []:
            if not isinstance(entry, dict):
                continue
            for f in entry.get("fileList", []) or []:
                if not isinstance(f, dict):
                    continue
                fname = f.get("showName") or f.get("fileName") or f.get("name")
                fid = f.get("fileURL") or f.get("fileId") or f.get("id")
                # 如果 fid 是完整的 URL，嘗試從其中擷取 /getfile/<id>
                if isinstance(fid, str) and fid.startswith("http"):
                    m = re.search(r"/getfile/([A-Za-z0-9]+)", fid)
                    if m:
                        fid = m.group(1)
                _add(fname, fid)
        return items

    # 回退到原先的通用處理
    if isinstance(file_list_json, list):
        candidates = file_list_json
    elif isinstance(file_list_json, dict):
        candidates = file_list_json.get("data") or file_list_json.get("files") or []
    else:
        candidates = []

    for x in candidates:
        if not isinstance(x, dict):
            continue
        fname = x.get("fileName") or x.get("filename") or x.get("name") or x.get("showName")
        fid = x.get("fileId") or x.get("fileID") or x.get("id") or x.get("fileURL")
        if isinstance(fid, str) and fid.startswith("http"):
            m = re.search(r"/getfile/([A-Za-z0-9]+)", fid)
            if m:
                fid = m.group(1)
        if fname and fid:
            items.append((str(fname), str(fid)))
    return items


def download_file(session: requests.Session, token: str, file_id: str, save_path: Path) -> None:
    # file_id 可能是純 id（例如 0900238e...）或是完整的 fileURL
    if isinstance(file_id, str) and file_id.startswith("http"):
        url = file_id
    else:
        url = DOWNLOAD_URL.format(file_id)

    last_err = None
    for attempt in range(1, RETRY + 1):
        try:
            with session.get(
                url,
                headers={"Authorization": f"Bearer {token}"},
                stream=True,
                timeout=TIMEOUT,
            ) as r:
                r.raise_for_status()
                with open(save_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=1024 * 128):
                        if chunk:
                            f.write(chunk)
            return
        except Exception as e:
            last_err = e
            time.sleep(0.8 * attempt)
    raise RuntimeError(f"下載失敗：{url} -> {save_path.name}，最後錯誤：{last_err}")


def read_input_file(path: str) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"找不到輸入檔：{path}")

    if p.suffix.lower() in [".xlsx", ".xls"]:
        return pd.read_excel(p)
    elif p.suffix.lower() == ".csv":
        return pd.read_csv(p)
    else:
        raise ValueError("輸入檔只支援 .xlsx/.xls/.csv")


def main():
    df = read_input_file(INPUT_FILE)

    if INPUT_COLUMN not in df.columns:
        raise ValueError(f"輸入檔沒有欄位：{INPUT_COLUMN}，目前欄位：{list(df.columns)}")

    pubnos = df[INPUT_COLUMN].dropna().astype(str).tolist()

    session = requests.Session()
    token = get_token(session)

    log_rows = []

    for pubno in tqdm(pubnos, desc="Downloading"):
        case_id = normalize_case_id(pubno)

        try:
            case_info = get_case_info(session, token, case_id)

            # token 過期時，有些系統會回特定 code/msg，你也可以在這裡加判斷後自動 refresh token
            case_no = case_info.get("caseNo") or case_info.get("caseNO")

            if not case_no:
                log_rows.append({"公開公告號": pubno, "caseId": case_id, "status": "FAIL", "reason": "查不到 caseNo"})
                continue

            file_list_json = get_file_list(session, token, str(case_no))
            file_items = parse_file_items(file_list_json)

            if not file_items:
                log_rows.append({"公開公告號": pubno, "caseId": case_id, "caseNo": case_no, "status": "FAIL", "reason": "沒有可下載檔案"})
                continue

            downloaded = 0
            for fname, fid in file_items:
                if not keyword_filter(fname):
                    continue

                out_name = safe_filename(f"{pubno}_{fname}")
                save_path = SAVE_DIR / out_name

                # 已存在就略過（可重跑）
                if save_path.exists() and save_path.stat().st_size > 0:
                    downloaded += 1
                    continue

                download_file(session, token, fid, save_path)
                downloaded += 1

            log_rows.append({"公開公告號": pubno, "caseId": case_id, "caseNo": case_no, "status": "OK", "downloaded_files": downloaded})

            time.sleep(SLEEP_BETWEEN_CASES)

        except Exception as e:
            log_rows.append({"公開公告號": pubno, "caseId": case_id, "status": "FAIL", "reason": str(e)})

    # 輸出 log
    log_df = pd.DataFrame(log_rows)
    log_path = Path("download_log.csv")
    try:
        log_df.to_csv(log_path, index=False, encoding="utf-8-sig")
        print(f"完成！下載資料夾：pdf_downloads/，log：{log_path}")
    except PermissionError:
        # 如果檔案被其它程式開啟（例如 Excel），則改寫入含時間戳的新檔名
        import datetime
        alt = Path(f"download_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        log_df.to_csv(alt, index=False, encoding="utf-8-sig")
        print(f"原始 log 檔案無法寫入（被其他程式鎖定），已改寫入：{alt}")


if __name__ == "__main__":
    main()
