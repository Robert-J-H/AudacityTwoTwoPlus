# -*- coding: utf-8 -*-
from __future__ import absolute_import, division, unicode_literals
# Audacity App Module for NVDA
# Copyright (c) 2017 Robert Hänggi and NVDA Access
import os, sys, api, ui, gui, re, speech, time, tones
impPath = os.path.abspath(os.path.dirname(__file__))
sys.path.append(impPath)
from builtins import zip, str, range
del sys.path[-1]
import appModuleHandler
from appModules import __path__ as paths
import addonHandler
from .appVars import *
import controlTypes
from comtypes import client
import IAccessibleHandler
import queueHandler
import inputCore
from keyboardHandler import KeyboardInputGesture as KIGesture
import NVDAObjects.IAccessible
import scriptHandler
import textInfos
from logHandler import log
from nvwave import playWaveFile
from wx import CallLater
from NVDAObjects.window import LiveText, DisplayModelLiveText
import winsound
SCRCAT_AUDACITY = _('Audacity')
MOUSEEVENTF_WHEEL=0x0800
dataPath=os.path.join(os.path.dirname(__file__).decode('mbcs'), 'data\\')
menuFull=[]
assignedShortcuts={}
toolBars={}
lastStatus='Stopped.'
suppressStatus=False
objCounter=0

def firstNum(string):
	for i, c in enumerate(string):
		if c.isdigit() and c!='0':
			return i

def getPipeFile():
	# could be used to employ a quasi-pipe to Audacity
	# in order to read e.g. Peaks or RMS of a selection
	# needs the "Screen Reader Support" Nyquist plug-in to be installed and enabled
	# in the analyze menu of Audacity.
	with open(os.environ['userprofile']+'\\audacity_scr_rdr_pipe.tmp') as file:
		data = file.read().splitlines()
		data = list([x.title() for x in data])
	return dict(list(zip(data[::2],data[1::2])))

class AppModule(appModuleHandler.AppModule):
	pastTime=currentTime=None
	deltaTime=[]
	tapCounter=0
	tapMedian=200.0
	navGestures={}
	outBox=CallLater(1, ui.message, None)

	def __init__(self, *args, **kwargs):
		super(AppModule, self).__init__(*args, **kwargs)
		self._audacityInputHelp=False
		self.helpPath=addonHandler.getCodeAddon().getDocFilePath()
		self.seenHandles=set()

	def _getByVersion(self, control):
		if int(''.join(x for x in self.productVersion[:-1] if x.isdigit()))>=220:
			return {'audio':11, 'start':14, 'length':15, 'center':16,'end':17}.get(control)
		else:
			return {'audio':14, 'start':11, 'length':12, 'end':12}.get(control)

	def speakAction(self, action):
		queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, action)

	def event_appModule_gainFocus(self):
		controlTypes.silentRolesOnFocus.add(controlTypes.ROLE_PANE)
		self.bindGesture('kb:applications', 'replaceApplications')
		inputCore.manager._captureFunc = self._inputCaptor

	def event_appModule_loseFocus(self):
		inputCore.manager._captureFunc = None
		controlTypes.silentRolesOnFocus.remove(controlTypes.ROLE_PANE)

	def _inputCaptor(self, gesture):
		global suppressStatus
		if api.getFocusObject().windowControlID != 1003 or lastStatus=='Recording.' or gesture.isNVDAModifierKey or gesture.isModifier:
		#if api.getForegroundObject().windowControlID != 68052 or lastStatus=='Recording.' or gesture.isNVDAModifierKey or gesture.isModifier:
			return True
		script = gesture.script
		lookup=assignedShortcuts.get(gesture.normalizedIdentifiers[1], None) 
		if self._audacityInputHelp:
			scriptName = scriptHandler.getScriptName(script) if script else ''
			if scriptName == 'toggleAudacityInputHelp':
				return True
			if lookup:
				queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, lookup[2]+'->'+lookup[0])
			if scriptName and not lookup:
				queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, scriptName)
		else:
			# don't speak "Stopped" for preview commands
			if lookup and lookup[0] in shouldNotReportStatus:
				suppressStatus=True
			else:
				suppressStatus=False
			if lookup and not lookup[0] in shouldNotAutoSpeak:
				queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, lookup[0])
		return not self._audacityInputHelp

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		try:
			name=obj.name
			role=obj.role
			windowText=obj.windowText
			windowControlID=obj.windowControlID
			childID = obj.IAccessibleChildID
		except:
			return
		#if (windowControlID in [2723]):
			#clsList.insert(0, SelectionControls)
		if (windowText=='Track Panel' and windowControlID==1003 and childID>=0):
				if ' Label Track' in name:
					clsList.insert(0, LabelTrack)
				else:
					clsList.insert(0, Track)

	def event_NVDAObject_init(self,obj):
		handle=obj.windowHandle
		if handle in self.seenHandles:
			pass
		if obj.windowClassName=='#32770' and obj.role==controlTypes.ROLE_PANE:
			#obj.role=controlTypes.ROLE_DIALOG
			obj.isFocusable=False
		# avoid the ampersand in dialogs
		if obj.windowClassName=='Button' and not obj.role in [controlTypes.ROLE_MENUBAR, controlTypes.ROLE_MENUITEM, controlTypes.ROLE_POPUPMENU]:
			obj.name = api.winUser.getWindowText(obj.windowHandle).replace('&','')
		# define the toolbars as real toolbars
		# such that the name is spoken automatically on focus entered.
		if obj.role==controlTypes.ROLE_PANE and obj.name: 
			if 		('audacity' in obj.name.lower()) and obj.firstChild.role!= controlTypes.ROLE_STATUSBAR:
				obj.role=controlTypes.ROLE_TOOLBAR
				obj.name=obj.name.lstrip('Audacity ').rstrip('Toolbar')
			elif obj.name=='Timeline':
				obj.role=controlTypes.ROLE_RULER
		# groupings for controls in e.g. preferences
		if obj  and obj.role in [5, 6, 8, 9, 13, 24, 36] \
			and not obj.windowHandle in self.seenHandles \
			and obj.container and not obj.container.role==controlTypes.ROLE_GROUPING:
			# Code snippet by David
			# work around for reading group boxes. In Audacity, any group box is a previous
			# sibling of a control, rather than a previous sibling of the parent of a control.
			# restrict the search two the roles listed above
			groupBox = IAccessibleHandler.findGroupboxObject(obj)
			if groupBox:
				obj.container = groupBox
			else:
				# don't check again
				self.seenHandles.add(obj.windowHandle)
		if obj.windowClassName=='#32768' and obj.role == controlTypes.ROLE_POPUPMENU:
			obj.name='DropDown'
		if obj.role==11 and '\\' in obj.name:
			# rearrange the items in the recent files menu such
			# that the file name comes before the full qualified path.
			obj.name='{2}, {0}{1}{2}'.format(*obj.name.rpartition('\\'))
		# append percent to slider positions
		if obj and obj.role==24 and obj.name==None and obj.previous and obj.previous.role==8:
			obj.description=' %'

	def replaceMulti(self, item, old, new):
		while len(old)!=0:
			item=str(item).replace(old.pop(0), new.pop(0))
		return item

	def _mapAudacityKeys(self, obj, parentName):
		cmd, _, shortcut=obj.name.partition('\x09')
		if bool(shortcut):
			shortcut=self.replaceMulti(shortcut, \
				[' ', 'Ctrl', 'Left', 'Right', 'Up', 'Down', 'Pageuparrow', 'Pagedownarrow', 'Return'], \
				['', 'control','leftarrow','rightarrow', 'uparrow', 'downarrow', 'pageup', 'pagedown', 'enter']).lower()
			ncmd=KIGesture.fromName(shortcut)
			cmd = cmd.title()
			if cmd in canditatesStartTime + canditatesEndTime + canditatesLength:
				self.navGestures[ncmd.identifiers[1]]=self.replaceMulti(cmd, [' ','(',')','/'], ['', '', '', ''])
			assignedShortcuts[ncmd.normalizedIdentifiers[1]]=(cmd, obj, parentName.title())

	def _get_Menus(self, obj):
		if len(assignedShortcuts)==0 and obj.previous and obj.previous.role==controlTypes.ROLE_MENUBAR:
			menus=obj.previous.children
			# 9 menus for V2.1.3 and 10(+2) for V2.2.0
			# 11 menus for V2.2.2 and presumably 12for V2.3.0
			if len(menus) >=9:
				signet=dataPath+'signet.wav'
				playWaveFile(signet)
				for menu in menus:
					queueHandler.queueFunction(queueHandler.eventQueue,self._get_menuItem,menu)
			return True

	# Sadly, recursion sucks in this context
	def _get_menuItem(self, obj):
		parentName=obj.name
		if obj.role==11:
			menuFull.append(obj)
			self._mapAudacityKeys(obj, parentName)
		if obj.firstChild:
			for level1 in obj.firstChild.children:
				if level1.role==11:
					menuFull.append( level1)
					self._mapAudacityKeys(level1, obj.name)
				if level1.firstChild:
					for level2 in level1.firstChild.children:
						if level2.role==11:
							menuFull.append(level2)
							self._mapAudacityKeys(level2, level1.name)
						if level2.firstChild:
							for level3 in level2.firstChild.children:
								if level3.role==11:
									menuFull.append(level3)
									self._mapAudacityKeys(level3, level2.name)

	def _update_Toolbars(self,callback=0):
		global toolBars
		obj=api.getForegroundObject()
		if not obj or obj.name == None:
			try:
				obj=api.getFocusAncestors()[2]
			except:
				obj=api.getFocusObject().parent.parent.parent
		if not obj:
			return
		winHandle=obj.windowHandle
		winClass=obj.windowClassName
		if winHandle in toolBars:
			return True
		else:
			activeToolBars={tb.name.lstrip('Audacity ').rstrip('Toolbar') : tb for tb in obj.recursiveDescendants if tb.role==controlTypes.ROLE_TOOLBAR}
			if bool(len(activeToolBars)):
				toolBars[winHandle]=activeToolBars
				return True
			elif callback!=1:
				self._update_Toolbars(1)
			else:
				return

	def getToolBar(self,key):
		obj=api.getForegroundObject()
		if obj.name==None:
			obj=api.getFocusAncestors()[2]
		wh=api.winUser.getForegroundWindow()
		if wh in toolBars:
			return toolBars[wh].get(key)
		elif obj.windowClassName=='wxWindowNR' and self._update_Toolbars():
			return toolBars[wh].get(key)

	def event_foreground(self, obj, nextHandler):
		speech.cancelSpeech()
		self._get_Menus(obj)
		nextHandler()

	def event_gainFocus(self, obj, nextHandler):
		# Nyquist effects with an unspoken unit and an unnamed slider 
		if obj and obj.role==8 and obj.windowControlID in range(12000,13000) \
			and obj.next and obj.next.role==24 and obj.next.name==None \
			and obj.next.next and obj.previous.location[0]!=obj.next.next.location[0]:
			try:
				obj.value+=' '+obj.next.next.name
			except:
				pass
		#Mainly for the Compressor effect
		if ((obj.role==8 and not controlTypes.STATE_READONLY in obj.states)or \
		obj.role==24 )and \
		obj.location[3]<obj.location[2] and \
		(obj.next and obj.next.name is not None) and \
		obj.next.role==7 and \
		(obj.next.next is None or \
		obj.next.name not in (obj.next.next.name or '')):
			try:
				if not obj.next.name in obj.name:
					obj.name+=' '+obj.next.name
			except:
				pass
		# suppress things like[Panel, Track View Table, TABLEROW...]
		if obj.windowText =='Track Panel':
			speech.cancelSpeech()
		if obj.windowClassName=='#32768' and obj.role == controlTypes.ROLE_POPUPMENU:
			# focus on first item for context menu
			name, obj.name=obj.name, None
			if name=='DropDown':
				KIGesture.fromName('downArrow').send()
		nextHandler()

	def event_valueChange(self, obj, nextHandler):
		if obj and obj.hasFocus and obj.role==24:
			#tones.beep(1000,10)
			if obj.name==None and obj.previous and obj.previous.role==8:
				try:
					# A bug in the Nyquist sliders, there's always a zero appended 
					val=str(eval(obj.previous.value))
					ui.message(val)
				except:
					pass
			elif not (obj.location[3]>obj.location[2] or obj.container.role==controlTypes.ROLE_TOOLBAR):
				val=obj.previous.value
				if val:
					obj.name=' '.join([obj._get_name(),val])
					ui.message(val)
				else:
					obj.name=obj.previous.name
					ui.message(obj.previous.name)
		nextHandler()

	def event_nameChange(self, obj, nextHandler):
		global lastStatus, suppressStatus
		name=obj.name
		if not suppressStatus and \
			name in ['Stopped.','Playing Paused.', 'Recording Paused.'] and \
			name !=lastStatus:
			ui.message(name.rstrip('.'))
		if name in ['Recording.','Playing.','Stopped.','Playing Paused.', 'Recording Paused.']:
			lastStatus=name
		nextHandler()

	def script_systemMixer(self,gesture):
		os.startfile('sndvol.exe')
	script_systemMixer.__doc__=_('Opens the System Volume Mixer, same as rightclicking on the speaker icon in the tray')
	script_systemMixer.category=SCRCAT_AUDACITY

	def script_replaceApplications(self, gesture):
		# This is necessary because the usage of the applications key deselects all tracks.
		KIGesture.fromName('shift+f10').send()

	def script_states(self,gesture):
		ui.browseableMessage(''.join((x+'\n' for x in toolBars.get(api.getForegroundObject().windowHandle).keys())),'')
		#ui.browseableMessage(repr(toolBars))
	script_states.__doc__=_('Reports all tool bars that are currently docked.')
	script_states.category=SCRCAT_AUDACITY

	def script_help(self, gesture):
		with open(self.helpPath,'r') as helpFile:
			help_html=helpFile.read()
			speech.cancelSpeech
			help_html=''.join(re.match('(.*<body>)(.*<hr />)(.*)',help_html,flags=re.DOTALL).groups()[::2])
			ui.browseableMessage(help_html,'Help',True)
	script_help.__doc__=_('Shows the accompanying help for the Audacity addon')
	script_help.category=SCRCAT_AUDACITY

	def script_guide(self, gesture):
		with open(dataPath+'Audacity 2.2.0 Guide.htm','r') as guide:
			guide=guide.read()
			speech.cancelSpeech
			ui.browseableMessage(guide,'Guide',True)
	script_guide.__doc__=_('Shows the famous JAWS Guide for Audacity as browseable document. All quick navigation keys allowed including find dialog and elements list.')
	script_guide.category=SCRCAT_AUDACITY

	def script_info(self,gesture):
		#ui.message(str(getPipeFile()['Channels']))
		out=''
		for x in menuFull:
			if x.role==11 and not controlTypes.STATE_HASPOPUP in x.states:
				out+=repr(x.name.partition('\x09')[0])+',\n'
		# ui.browseableMessage(out)
		ui.browseableMessage('\n'.join((api.getCaretObject().name,api.getFocusObject().name,str(api.getForegroundObject().windowHandle),str(api.winUser.getForegroundWindow()))))
	script_info.__doc__=_('debugging info')
	script_info.category=SCRCAT_AUDACITY

	def _paste_safe(self, text, obj=None, label=False):
		try:
			temp=api.getClipData()
		except:
			temp=''
		api.copyToClip(text)
		if label==True:
			KIGesture.fromName('p').send()
			KIGesture.fromName('control+alt+v').send()
		else:
			KIGesture.fromName('control+v').send()
		api.processPendingEvents()
		if obj:
			obj.reportFocus()
		api.copyToClip(temp)

	def script_announcePlaybackPeak(self,gesture):
		pPeak=self.getToolBar('Playback Meter ').getChild(1).name.partition('Peak')[2] 
		ui.message(pPeak)
	script_announcePlaybackPeak.__doc__=_('Reports the current playback level.')
	script_announcePlaybackPeak.category=SCRCAT_AUDACITY

	def script_announceRecordingPeak(self,gesture):
		repeatCount=scriptHandler.getLastScriptRepeatCount()
		#if repeatCount==0:
		rPeak=self.getToolBar('Recording Meter ').getChild(1).name.partition('Peak')[2] 
		ui.message(rPeak)
	script_announceRecordingPeak.__doc__=_('Reports the current recording level.')
	script_announceRecordingPeak.category=SCRCAT_AUDACITY

	def script_announceTempo(self,gesture):
		repeatCount=scriptHandler.getLastScriptRepeatCount()
		obj=api.getFocusObject()
		if repeatCount==0:
			if obj.role==controlTypes.ROLE_EDITABLETEXT:
				text=str(round(60.0/self.tapMedian,1))
				self._paste_safe(text, obj)
			else:
				ui.message(str(round(60.0/self.tapMedian,1))+' bpm')
		elif repeatCount==1:
			tempo=(60.0/self.tapMedian)
			text='Tempo not in classical range'
			for i in range(0, len(self.__tempi)-2,2):
				if tempo>self.__tempi[i] and tempo<=self.__tempi[i+2]:
					text=self.__tempi[i+1]
			ui.message(text)
	script_announceTempo.__doc__=_('''Reports the tempo in beats per minute for the last tapping along.
	Replaces the selection in an edit box if pressed once.
	Pressed twice, the tempo is given as classical description.''')
	script_announceTempo.category=SCRCAT_AUDACITY

	def script_tempoTapping(self,gesture):
		self.currentTime=time.clock()
		if self.pastTime==None or self.currentTime-self.pastTime>=2:
			tones.beep(4500, 12, 0, 50)
			self.pastTime=time.clock()
			self.tapCounter=0
			self.deltaTime=[]
		else:
			tones.beep(3000, 8, 0, 50)
			self.deltaTime.insert(0, self.currentTime-self.pastTime)
			self.pastTime=self.currentTime
			if len(self.deltaTime)>=8:
				outlayer=max([abs(x-self.tapMedian) for x in self.deltaTime])
				try:
					self.deltaTime.remove(self.tapMedian+outlayer)
				except:
					self.deltaTime.remove(self.tapMedian-outlayer)
			self.tapMedian=sum(self.deltaTime)/len(self.deltaTime)
			self.tapCounter=(self.tapCounter+1)%4
	script_tempoTapping.__doc__=_('use this key to tap along for some measures. The found tempo can be read out by the announce temp shortcut (normally NVDA+pause).')
	script_tempoTapping.category=SCRCAT_AUDACITY

	def script_reportColumn(self,gesture):
		obj=api.getFocusObject()
		try:
			info=obj.makeTextInfo(textInfos.POSITION_CARET)
			pos=info.bookmark.startOffset
			info.expand(textInfos.UNIT_LINE)
			info.collapse()
			speech.speak(str(pos-info.bookmark.startOffset))
		except:
			pass
	script_reportColumn.__doc__=_('Reports the Column in a edit box where the cursor is situated. Useful for the Nyquist Prompt')
	script_reportColumn.category=SCRCAT_AUDACITY

	def script_wheelForward(self, gesture):
		api.winUser.mouse_event(MOUSEEVENTF_WHEEL,0,0,120,None)
	script_wheelForward.__doc__=_('Simulates a turn of the Mouse wheel forward.')
	script_wheelForward.category=SCRCAT_AUDACITY

	def script_wheelBack(self, gesture):
		api.winUser.mouse_event(MOUSEEVENTF_WHEEL,0,0,-120,None)
	script_wheelBack.__doc__=_('Simulates a turn of the Mouse wheel back.')
	script_wheelBack.category=SCRCAT_AUDACITY

	def script_toggleAudacityInputHelp(self,gesture):
		if not lastStatus in ['Playing Paused.', 'Recording Paused.', 'Stopped.']:
			return
		elif api.getFocusObject().windowControlID != 1003:
			ui.message("Input Help is only available in the Track View")
			return
		self._audacityInputHelp=not self._audacityInputHelp
		stateOn = 'Audacity input help on'
		stateOff = 'Audacity input help off'
		state = stateOn if self._audacityInputHelp else stateOff
		ui.message(state)
	script_toggleAudacityInputHelp.__doc__=_('Turns Audacity input help on or off. When on, any input such as pressing a key on the keyboard will tell you what command  is associated with that input, if any')
	script_toggleAudacityInputHelp.category=SCRCAT_AUDACITY

	def getTime(self,child):
		rslt=self.getToolBar('Selection ').getChild(child).name.replace(',','')
		rslt=rslt.replace('+', ' ')
		for index, c in enumerate(rslt):
			if c.isdigit(): 
				break
		rslt=rslt[index :].split()
		for k in range(0, len(rslt), 2):
			val=float(rslt[k])
			if val==0:
				if rslt[k+1] in ['s', 'seconds', 'samples', '%']:
					minRslt=['0', rslt[k+1]]
				rslt[k]=rslt[k+1]=''
			elif val==int(val):
				rslt[k]=str(int(val))
			else:
				rslt[k]=str(val)
		for part in rslt[:]:
			if part=='':
				rslt.remove(part)
		rslt=' '.join(rslt)
		if rslt.isspace()or rslt=='':
			rslt=' '+' '.join(minRslt)+' '
		else:
			rslt=' '+rslt+' '
		for val in ['1 hour', '1 minute', '1 second']:
			rslt=rslt.replace(' '+val[0:3]+' ', ' '+val+', ')
		for val in ['hours, ', 'minutes, ', 'seconds']:
			rslt=rslt.replace(' '+val[0]+' ', ' '+val)
		return rslt

	__gestures={
		'kb:pause':'tempoTapping',
		'kb:f12':'toggleAudacityInputHelp',
		'kb:nvda+pause':'announceTempo',
		'kb:nvda+i':'reportColumn',
		'kb:nvda+g':'guide',
		'kb:nvda+h':'help',
		'kb:F8':'states',
		'kb:control+f8':'info',
		'kb:F9':'announcePlaybackPeak',
		'kb:F10':'announceRecordingPeak',
		'kb:Nvda+PageDown':'wheelBack',
		'kb:Nvda+pageUp':'wheelForward',
		'kb:Nvda+x':'systemMixer',
	}

	__tempi= [
		30, 'Grave', 42, 'Larghissimo', 44, 'Largo',
		47, 'Larghetto', 49, 'Lento', 54, 'Adagio',
		63, 'Adagieto', 69, 'Andante', 76, 'Andantino',
		84, 'Maestoso', 92, 'Moderato', 104, 'Allegretto',
		116, 'Animato', 126, 'Allegro', 138, 'Allegro Assai',
		152, 'Vivace', 176, 'Vivacissimo', 182, 'Presto',
		200, 'Prestissimo', 208, 'Presto prestissimo', 230
	]

class SelectionControls (NVDAObjects.IAccessible.IAccessible):
	speech.cancelSpeech()
	#self.displayText=''
	#self.name=''

	def __init__(self, *args, **kwargs):
		super(SelectionControl, self).__init__(*args, **kwargs)

	def script_changeAndPreview(self, gesture):
		gesture.send()
		api.getFocusObject().displayText=''
		api.getFocusObject().name==''
		KIGesture.fromName('Shift+F6').send()
	script_changeAndPreview.__doc__=_('When in a selection control, the value will be changed and a preview played.')
	script_changeAndPreview.category=SCRCAT_AUDACITY

	__gestures={
		'kb:shift+upArrow':'changeAndPreview',
		'kb:shift+downArrow':'changeAndPreview',
	}

class Track (NVDAObjects.IAccessible.IAccessible, AppModule):
	# prevent the roles table and row (= track panel and audio track) from being spoken 
	controlTypes.silentRolesOnFocus.add(controlTypes.ROLE_TABLEROW)
	controlTypes.silentValuesForRoles.add(controlTypes.ROLE_TABLEROW)
	controlTypes.silentRolesOnFocus.add(controlTypes.ROLE_TABLE)
	shouldAllowIAccessibleFocusEvent=True

	def initOverlayClass(self):
		self.appName=self.appModuleName

	def _get_next(self):
			if not self.positionInfo: return
			totalGroup,posGroup=self.positionInfo.values()
			if 0<posGroup<totalGroup:
				return self.parent.getChild(posGroup)

	def _get_previous(self):
			if not self.positionInfo: return
			totalGroup,posGroup=self.positionInfo.values()
			if 1<posGroup<=totalGroup:
				return self.parent.getChild(posGroup-2)

	def event_gainFocus(self):
		super(Track,self).event_gainFocus()
		if self.IAccessibleChildID>=0:
			self._update_Toolbars()
			self.bindGestures(self.navGestures)

	def script_quickMarker(self, gesture):
		KIGesture.fromName('control+m').send()
		KIGesture.fromName('enter').send()
	script_quickMarker.__doc__=_('Inserts a point label which is automatically closed')
	script_quickMarker.category=SCRCAT_AUDACITY


	def script_announceAudioPosition(self,gesture):
		audioPosition=self.getTime(self._getByVersion('audio'))
		repeatCount=scriptHandler.getLastScriptRepeatCount()
		if repeatCount==0:
			ui.message(audioPosition)
		elif repeatCount==1:
			speech.speakSpelling(audioPosition)
	script_announceAudioPosition.__doc__=_('Reports the audio position, both during Playback and while in stop mode.')
	script_announceAudioPosition.category=SCRCAT_AUDACITY

	def script_announceStart(self, gesture):
		repeatCount=scriptHandler.getLastScriptRepeatCount()
		if repeatCount==0:
			start = self.getTime(self._getByVersion('start'))+' Start'
		else:
			start = self.getTime(self._getByVersion('center'))+' Center'
		ui.message(start)
	script_announceStart.__doc__=_('Reports the cursor position or left selection boundary.')
	script_announceStart.category=SCRCAT_AUDACITY

	def script_announceEnd(self, gesture):
		repeatCount=scriptHandler.getLastScriptRepeatCount()
		if repeatCount==0:
			end = self.getTime(self._getByVersion('end'))+' end'
		else:
			end = self.getTime(self._getByVersion('length'))+' long'
		ui.message(end)
	script_announceEnd.__doc__=_('Reports the right selection boundary if end is chosen in the selection toolbar. Otherwise, the selection length is reported, followed by the word long.')
	script_announceEnd.category=SCRCAT_AUDACITY
	
	def script_reportSelectedTracks(self,gesture):
		repeatCount=scriptHandler.getLastScriptRepeatCount()
		selChoice=''
		selTracks=[]
		if repeatCount==0:
			selTracks=['Selected ']
			selChoice='Select On'
		elif repeatCount==1:
			selTracks=['Muted ']
			selChoice='Mute On'
		elif repeatCount==2:
			selTracks=['Soloed ']
			selChoice='Solo On'
		for track in self.parent.children:
			if selChoice in track.name:
				selTracks.append(track.name.replace(selChoice, ''))
				try:
					track.doAction(1)
				except NotImplementedError:
					pass
		if len(selTracks)==1:
			ui.message('No Tracks '+selTracks[0])
		else:
			selTracks[0]+='is: ' if len(selTracks)==2 else ' are:'
			ui.message(', '.join(selTracks)) 
	script_reportSelectedTracks.__doc__=_('Reports the currently selected tracks, if any. Reports muted tracks if pressed twice or soloed tracks if pressed three times.')
	script_reportSelectedTracks.category=SCRCAT_AUDACITY

	def script_pageUpByThree(self, gesture):
		for i in range(3):
			KIGesture.fromName('upArrow').send()
	script_pageUpByThree.__doc__=_('Moves three tracks up.')
	script_pageUpByThree.category=SCRCAT_AUDACITY

	def script_pageDownByThree(self, gesture):
		for i in range(3):
			KIGesture.fromName('downArrow').send()
	script_pageDownByThree.__doc__=_('Moves three Tracks down.')
	script_pageDownByThree.category=SCRCAT_AUDACITY

	def script_expandLeft(self, gesture):
		KIGesture.fromName('Shift+LeftArrow').send()
		KIGesture.fromName('Shift+F6').send()
	script_expandLeft.__doc__=_('Expands the selection at the left boundary and plays a preview.')
	script_expandLeft.category=SCRCAT_AUDACITY

	def script_reduceLeft(self, gesture):
		KIGesture.fromName('Control+Shift+RightArrow').send()
		KIGesture.fromName('Shift+F6').send()
	script_reduceLeft.__doc__=_('Contracts the selection at the left boundary and plays a preview.')
	script_reduceLeft.category=SCRCAT_AUDACITY

	def script_expandRight(self, gesture):
		KIGesture.fromName('Shift+RightArrow').send()
		KIGesture.fromName('Shift+F8').send()
	script_expandRight.__doc__=_('Expands the selection at the right boundary and plays a preview.')
	script_expandRight.category=SCRCAT_AUDACITY

	def script_reduceRight(self, gesture):
		KIGesture.fromName('Control+Shift+LeftArrow').send()
		KIGesture.fromName('Shift+F8').send()
	script_reduceRight.__doc__=_('Contracts the selection at the right boundary and plays a preview.')
	script_reduceRight.category=SCRCAT_AUDACITY

	def autoTime(self, gesture, which=0):
		def pick(control):
			repeatCount=scriptHandler.getLastScriptRepeatCount()
			oldTimes={x:self.getTime(self._getByVersion(x)) for x in {control, 'start', 'end', 'length'}}
			gesture.send()
			newTimes={x:self.getTime(self._getByVersion(x)) for x in {control, 'start', 'end', 'length'}}
			if newTimes == oldTimes:
				winsound.MessageBeep()
			else:
				newTimes['audio']=self.getTime(self._getByVersion('audio'))
				message= newTimes[control]+' selected' if control=='length' else newTimes[control]
				if repeatCount==0:
					if self.outBox and self.outBox.running:
						for pending in self.outBox._CallLater__RUNNING-set((self.outBox,)):
							pending.Stop() 
					else:
						ui.message(message)
				else:
					self.outBox.Restart(60, message)
		# Do nothing during recording
		if lastStatus=='Recording.':
			gesture.send()
		elif lastStatus=='Playing.':
			gesture.send()
			if which in [2,4]:
				queueHandler.queueFunction(queueHandler.eventQueue, playWaveFile, dataPath+'selection_start.wav')
			elif which in [3,5]:
				queueHandler.queueFunction(queueHandler.eventQueue, playWaveFile, dataPath+'selection_end.wav')
			else:
				return
		elif lastStatus in ['Stopped.', 'Playing Paused.', 'blank', None]:
			if which in [0,2]:
				pick('audio')
			elif which in [4,5]:
				pick('length')
			elif which in [1,3]:
				if 'Paused' in lastStatus:
					pick('audio')
				else:
					pick('end')

	# commands that should speak the new selection length
	def script_SelectionToStart(self, gesture):
		self.autoTime(gesture,4)
	def script_SelectionToEnd(self, gesture):
		self.autoTime(gesture,5)
	def script_TrackStartToCursor(self, gesture):
		self.autoTime(gesture, 4)
	def script_CursorToTrackEnd(self, gesture):
		self.autoTime(gesture, 5)
	#commands that should speak the start or end time of the selection
	def script_AddLabelAtPlaybackPosition(self, gesture):
		self.autoTime(gesture)
	def script_PlayStopAndSetCursor(self, gesture):
		self.autoTime(gesture)
	def script_SelectionStart(self, gesture):
		self.autoTime(gesture)
	def script_SelectionEnd(self, gesture):
		self.autoTime(gesture)
	def script_TrackStart(self, gesture):
		self.autoTime(gesture)
	def script_TrackEnd(self, gesture):
		self.autoTime(gesture)
	def script_PreviousClipBoundary(self, gesture):
		self.autoTime(gesture)
	def script_NextClipBoundary(self, gesture):
		self.autoTime(gesture)
	def script_ProjectStart(self, gesture):
		self.autoTime(gesture)
	def script_ProjectEnd(self, gesture):
		self.autoTime(gesture)
	def script_ShortSeekLeftDuringPlayback(self, gesture):
		self.autoTime(gesture)
	def script_LongSeekLeftDuringPlayback(self, gesture):
		self.autoTime(gesture)
	def script_SelectionExtendLeft(self, gesture):
		self.autoTime(gesture)
	def script_SetOrExtendLeftSelection(self, gesture):
		self.autoTime(gesture,2)
	def script_LeftAtPlaybackPosition(self, gesture):
		self.autoTime(gesture,2)
	def script_SelectionContractLeft(self, gesture):
		self.autoTime(gesture)
	def script_CursorLeft(self, gesture):
		self.autoTime(gesture)
	def script_CursorShortJumpLeft(self, gesture):
		self.autoTime(gesture)
	def script_CursorLongJumpLeft(self, gesture):
		self.autoTime(gesture)
	def script_ClipLeft(self, gesture):
		self.autoTime(gesture)
	def script_ClipRight(self, gesture):
		self.autoTime(gesture)
	def script_ShortSeekRightDuringPlayback(self, gesture):
		self.autoTime(gesture)
	def script_CursorRight(self, gesture):
		self.autoTime(gesture)
	def script_SelectionExtendRight(self, gesture):
		self.autoTime(gesture,1)
	def script_SetOrExtendRightSelection(self, gesture):
		self.autoTime(gesture,3)
	def script_RightAtPlaybackPosition(self, gesture):
		self.autoTime(gesture,3)
	def script_SelectionContractRight(self, gesture):
		self.autoTime(gesture,1)
	def script_LongSeekRightDuringPlayback(self, gesture):
		self.autoTime(gesture,1)
	def script_CursorShortJumpRight(self, gesture):
		self.autoTime(gesture,1)
	def script_CursorLongJumpRight(self, gesture):
		self.autoTime(gesture,1)


	__gestures={
		'kb:m':'quickMarker',
		'kb:nvda+a':'announceAudioPosition',
		'kb:nvda+j':'announceStart',
		'kb:nvda+k':'announceEnd',
		'kb:pageup':'pageUpByThree',
		'kb:pagedown':'pageDownByThree',
		'kb:nvda+e':'reportSelectedTracks',
		'kb:nvda+volumeDown':'expandLeft',
		'kb:nvda+volumeUp':'reduceLeft',
		'kb:nvda+shift+volumeUp':'expandRight',
		'kb:nvda+shift+volumeDown':'reduceRight',
	}

class LabelTrack(DisplayModelLiveText):
	shouldAllowIAccessibleFocusEvent=True
	editMode=0
	navMode=0

	def initOverlayClass(self):
		self.isFocusable=True

	def event_gainFocus(self):
		super(LabelTrack,self).event_gainFocus()
		self.startMonitoring()

	def event_loseFocus(self):
		self.stopMonitoring()

		def _getTextLines(self):
			return self.displayText.split()
	def event_textChange(self):
		super(LabelTrack, self).event_textChange()
		# tones.beep(900,10)