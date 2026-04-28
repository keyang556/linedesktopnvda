from __future__ import annotations

import ast
import re
from pathlib import Path


def _load_image_menu_helpers():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	namespace = {"re": re}
	needed_assignments = {
		"_CJK_CHAR",
		"_CJK_SPACE_RE",
		"_IMAGE_ATTACHMENT_MENU_KEYWORDS",
		"_STICKER_MESSAGE_MENU_KEYWORDS",
	}
	needed_functions = {
		"_removeCJKSpaces",
		"_extractDownloadDeadlineAnnouncement",
		"_looksLikeImageAttachmentMenu",
		"_looksLikeStickerMessageMenu",
	}

	for node in module.body:
		if isinstance(node, ast.Assign):
			names = {target.id for target in node.targets if isinstance(target, ast.Name)}
			if names & needed_assignments:
				exec(
					compile(
						ast.Module(body=[node], type_ignores=[]),
						str(module_path),
						"exec",
					),
					namespace,
				)
		elif isinstance(node, ast.FunctionDef) and node.name in needed_functions:
			exec(
				compile(
					ast.Module(body=[node], type_ignores=[]),
					str(module_path),
					"exec",
				),
				namespace,
			)

	return namespace


helpers = _load_image_menu_helpers()


def _load_find_copy_menu_item():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8-sig")
	module = ast.parse(source)
	for node in module.body:
		if isinstance(node, ast.FunctionDef) and node.name == "_copyAndReadMessage":
			for child in node.body:
				if isinstance(child, ast.FunctionDef) and child.name == "_findCopyMenuItem":
					return child
	raise AssertionError("Could not find _findCopyMenuItem in line.py")


def _target_has_name(target, name):
	return any(isinstance(node, ast.Name) and node.id == name for node in ast.walk(target))


class _UnderscoreAssignmentVisitor(ast.NodeVisitor):
	def __init__(self):
		self.bad_lines = []

	def visit_Assign(self, node):
		if any(_target_has_name(target, "_") for target in node.targets):
			self.bad_lines.append(node.lineno)
		self.visit(node.value)

	def visit_AnnAssign(self, node):
		if _target_has_name(node.target, "_"):
			self.bad_lines.append(node.lineno)
		if node.value is not None:
			self.visit(node.value)

	def visit_AugAssign(self, node):
		if _target_has_name(node.target, "_"):
			self.bad_lines.append(node.lineno)
		self.visit(node.value)

	def visit_NamedExpr(self, node):
		if _target_has_name(node.target, "_"):
			self.bad_lines.append(node.lineno)
		self.visit(node.value)

	def visit_For(self, node):
		if _target_has_name(node.target, "_"):
			self.bad_lines.append(node.lineno)
		for stmt in node.body:
			self.visit(stmt)
		for stmt in node.orelse:
			self.visit(stmt)

	def visit_AsyncFor(self, node):
		self.visit_For(node)

	def visit_With(self, node):
		for item in node.items:
			if item.optional_vars is not None and _target_has_name(item.optional_vars, "_"):
				self.bad_lines.append(node.lineno)
		for stmt in node.body:
			self.visit(stmt)

	def visit_AsyncWith(self, node):
		self.visit_With(node)

	def visit_ExceptHandler(self, node):
		if node.name == "_":
			self.bad_lines.append(node.lineno)
		for stmt in node.body:
			self.visit(stmt)


def test_image_attachment_menu_is_detected_from_photo_actions():
	assert helpers["_looksLikeImageAttachmentMenu"](
		"回覆\n分享\n刪除\n轉為文字\n掃描行動條碼\n另存新檔\n"
		"傳送至 Keep 筆記\n儲存至記事本\n新增至相簿\n設為聊天室背景",
	)


def test_self_sent_message_menu_is_not_treated_as_image_attachment():
	assert not helpers["_looksLikeImageAttachmentMenu"](
		"回覆\n分享\n收回\n刪除\n翻譯\n傳送至 Keep 筆記\n儲存至記事本\n設為公告",
	)


def test_generic_save_as_menu_without_photo_actions_is_not_treated_as_image_attachment():
	assert not helpers["_looksLikeImageAttachmentMenu"](
		"回覆\n另存新檔\n傳送至 Keep 筆記\n儲存至記事本",
	)


def test_sticker_message_menu_is_detected_from_sticker_shop_action():
	assert helpers["_looksLikeStickerMessageMenu"]("回覆\n刪除\n貼圖小舖")
	assert helpers["_looksLikeStickerMessageMenu"]("回覆\n刪除\n貼 圖 小 舖")


def test_non_sticker_menu_is_not_treated_as_sticker_message():
	assert not helpers["_looksLikeStickerMessageMenu"](
		"回覆\n分享\n刪除\n轉為文字\n掃描行動條碼\n另存新檔",
	)


def test_extract_download_deadline_from_same_line():
	assert (
		helpers["_extractDownloadDeadlineAnnouncement"]("檔案\n下載期限：2026/05/01")
		== "下載期限：2026/05/01"
	)


def test_extract_download_deadline_from_following_line():
	assert (
		helpers["_extractDownloadDeadlineAnnouncement"]("下載期限\n2026 年 5 月 1 日")
		== "下載期限：2026 年 5 月 1 日"
	)


def test_find_copy_menu_item_does_not_shadow_translation_callable():
	visitor = _UnderscoreAssignmentVisitor()
	visitor.visit(_load_find_copy_menu_item())
	assert not visitor.bad_lines
