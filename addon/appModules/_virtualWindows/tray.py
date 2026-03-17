from .._virtualWindow import VirtualWindow
from controlTypes.role import Role
from controlTypes.state import State

class Tray(VirtualWindow):
	title = f'Line {Role.MENU.displayString}'
	
	@staticmethod
	def isMatchLineScreen(obj):
		try:
			return 'LcContextMenu' in obj.UIAElement.CurrentClassName
		except AttributeError:
			return False
		
	
	def makeElements(self):
		try:
			obj = self.obj.firstChild.next.lastChild
			self.elements.append({
				'name': '結束應用程式',
				'role': None,
				'clickPoint': self.rectGetCenterPoint(obj.location)
			})
			stateText = State.UNAVAILABLE.displayString if State.UNAVAILABLE in obj.parent.firstChild.states else ''
			isLoggedIn = not stateText or State.UNAVAILABLE in obj.previous.states
			if isLoggedIn:
				obj = obj.previous
				self.elements.append({
					'name': '登出' + (f' ({stateText})' if stateText else ''),
					'role': None,
					'clickPoint': self.rectGetCenterPoint(obj.location)
				})
			
			obj = obj.previous.previous
			self.elements.append({
				'name': '確認有無最新版本',
				'role': None,
				'clickPoint': self.rectGetCenterPoint(obj.location)
			})
			obj = obj.previous
			self.elements.append({
				'name': '關於LINE',
				'role': None,
				'clickPoint': self.rectGetCenterPoint(obj.location)
			})
			obj = obj.previous
			
			self.elements.append({
				'name': 'Keep筆記' + (f' ({stateText})' if stateText else ''),
				'role': None,
				'clickPoint': self.rectGetCenterPoint(obj.location)
			})
			obj = obj.previous
			self.elements.append({
				'name': '設定',
				'role': None,
				'clickPoint': self.rectGetCenterPoint(obj.location)
			})
			if not isLoggedIn:
				obj = obj.previous.previous
				self.elements.append({
					'name': '登入',
					'role': None,
					'clickPoint': self.rectGetCenterPoint(obj.location)
				})
			
			obj = obj.previous.previous
			self.elements.append({
				'name': '好友名單' + (f' ({stateText})' if stateText else ''),
				'role': None,
				'clickPoint': self.rectGetCenterPoint(obj.location)
			})
		finally:
			self.elements.reverse()
		
	
