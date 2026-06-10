"""Centralised multi-language label table for matching LINE's UI text.

The add-on detects menu items, dialog buttons and status text by matching the
strings LINE renders (via UIA names or OCR). Those strings were previously
hard-coded as Traditional Chinese literals scattered across the code base, so
the add-on only worked when LINE's UI language was Traditional Chinese.

This module collects them in one place keyed by a *canonical* string. The
canonical string is deliberately the Traditional Chinese label, because the
rest of the code dispatches on those exact literals (e.g.
``actionName == "收回"``); returning the canonical keeps every caller working
unchanged while letting OCR/UIA text in other languages still match.

Each entry maps the canonical label to the set of strings that should match it
across languages, including common OCR mis-reads. Languages that have not been
verified yet are simply omitted — add them here as they are confirmed; nothing
else needs to change.
"""

import difflib
import re


# canonical (zh-Hant) -> tuple of accepted strings across languages + OCR mis-reads.
# Keep the canonical string itself in the tuple.
_MESSAGE_MENU_LABELS = {
	"回覆": ("回覆", "回復", "回覧", "回复", "Reply", "返信", "ตอบกลับ"),
	"複製": ("複製", "复制", "複裂", "Copy", "コピー", "คัดลอก"),
	"分享": ("分享", "Share", "共有", "แชร์"),
	"刪除": ("刪除", "删除", "Delete", "削除", "ลบ"),
	"收回": ("收回", "Unsend", "送信取り消し", "取り消し", "ยกเลิกการส่ง"),
	"翻譯": ("翻譯", "翻译", "Translate", "翻訳", "แปล"),
	"傳送至Keep筆記": (
		"傳送至Keep筆記",
		"傳送至 Keep 筆記",
		"傳送至Keep",
		"傳送至 Keep",
		"傅送至Keep筆記",
		"Keep",
		"Save to Keep",
		"Keepに保存",
	),
	"儲存至記事本": ("儲存至記事本", "Save to Notes", "ノートに保存"),
	"設為公告": ("設為公告", "Pin", "Announce", "アナウンス", "ปักหมุด"),
	"另存新檔": ("另存新檔", "Save As", "Save as", "名前を付けて保存", "บันทึกเป็น"),
	"轉傳": ("轉傳", "Forward", "転送", "ส่งต่อ"),
	"貼圖小舖": ("貼圖小舖", "貼圖小鋪", "Sticker Shop", "スタンプショップ"),
	"轉為文字": ("轉為文字", "Convert to text", "テキストに変換", "แปลงเป็นข้อความ"),
	"掃描行動條碼": ("掃描行動條碼", "Scan QR code", "QRコードをスキャン"),
	"新增至相簿": ("新增至相簿", "Add to album", "アルバムに追加"),
	"設為聊天室背景": ("設為聊天室背景", "Set as chat background", "背景に設定"),
}

_MORE_OPTIONS_LABELS = {
	"開啟提醒": ("開啟提醒", "Turn on notifications", "通知をオン", "เปิดการแจ้งเตือน"),
	"關閉提醒": ("關閉提醒", "Turn off notifications", "通知をオフ", "ปิดการแจ้งเตือน"),
	"邀請": ("邀請", "Invite", "招待", "เชิญ"),
	"相簿": ("相簿", "Albums", "Album", "アルバム", "อัลบั้ม"),
	"照片・影片": (
		"照片・影片",
		"照片影片",
		"照片影⽚",
		"照片影像",
		"照片•影片",
		"照片‧影片",
		"眧片影片",
		"照片 影片",
		"Photos & Videos",
		"写真・動画",
	),
	"檔案": ("檔案", "Files", "File", "ファイル", "ไฟล์"),
	"連結": ("連結", "Links", "Link", "リンク", "ลิงก์"),
	"投票": ("投票", "訍疋", "Vote", "Poll", "投票", "โหวต"),
	"儲存聊天": ("儲存聊天", "Save chat", "トーク保存", "บันทึกแชท"),
	"背景設定": ("背景設定", "冃景言殳定", "背景设定", "Background", "背景設定"),
	"檢舉": ("檢舉", "Report", "通報", "รายงาน"),
	"退出群組": ("退出群組", "退出群组", "Leave group", "グループを退会", "ออกจากกลุ่ม"),
	"封鎖": ("封鎖", "Block", "ブロック", "บล็อก"),
}

# Incoming/outgoing call buttons. Used to locate buttons by UIA name or OCR.
_CALL_ANSWER = ("接聽", "接受", "accept", "answer", "応答", "รับสาย")
_CALL_DECLINE = ("拒絕", "decline", "reject", "拒否", "ปฏิเสธ")

# Recall confirmation dialog actions.
_RECALL_ACTIONS = {
	"收回": ("收回", "Unsend", "送信取り消し", "ยกเลิกการส่ง"),
	"取消": ("取消", "Cancel", "キャンセル", "ยกเลิก"),
	"無痕收回": ("無痕收回", "Unsend silently", "こっそり取り消し"),
}

# Photo-to-text first-run consent dialog actions.
_PHOTO_CONSENT_ACTIONS = {
	"同意": ("同意", "Agree", "同意する", "ยอมรับ", "ยินยอม"),
	"不同意": ("不同意", "Disagree", "同意しない", "ไม่ยอมรับ"),
}

# Microphone status keywords (the LINE control bar shows the toggle action,
# so "mute" means the mic is currently ON, and vice-versa).
_MIC_CURRENTLY_ON = ("關麥克風", "關閉麥克風", "Mute", "mute", "ミュート")
_MIC_CURRENTLY_OFF = ("開麥克風", "開啟麥克風", "Unmute", "unmute", "ミュート解除")


def _normalize(text):
	"""Collapse whitespace so spacing differences don't defeat matching."""
	return re.sub(r"\s+", "", (text or "").strip())


def allAliases(table, key):
	"""Return the accepted strings for one canonical key, or ()."""
	return table.get(key, ())


def matchLabel(text, table, threshold=0.62):
	"""Return the canonical key whose aliases best match ``text``.

	First tries substring containment against every alias (cheap and exact),
	then falls back to a fuzzy ratio against the canonical keys. Returns the
	canonical (Traditional Chinese) string so existing dispatch code that
	compares against those literals keeps working, or ``None``.
	"""
	normalized = _normalize(text)
	if not normalized:
		return None

	for canonical, aliases in table.items():
		for alias in aliases:
			if _normalize(alias) in normalized:
				return canonical

	bestKey = None
	bestRatio = 0.0
	for canonical in table:
		ratio = difflib.SequenceMatcher(None, normalized, _normalize(canonical)).ratio()
		if ratio > bestRatio:
			bestRatio = ratio
			bestKey = canonical
	if bestKey and bestRatio >= threshold:
		return bestKey
	return None


def containsAny(text, aliases):
	"""True if ``text`` contains any of the given alias strings (normalized)."""
	normalized = _normalize(text)
	if not normalized:
		return False
	return any(_normalize(alias) in normalized for alias in aliases)
