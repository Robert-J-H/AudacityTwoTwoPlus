# -*- coding: utf-8 -*-
from __future__ import absolute_import, division, unicode_literals
# Audacity App Module for NVDA
# Copyright (c) 2017 Robert Hänggi and NVDA Access
import os
import sys
impPath = os.path.abspath(os.path.dirname(__file__))
sys.path.append(impPath)
from builtins import zip, str, range
del sys.path[-1]
import appModuleHandler
from appModules import __path__ as paths
from .appVars import *
import api
import ui
import characterProcessing
import controlTypes
from comtypes import client
import IAccessibleHandler
import eventHandler
import queueHandler
import gui
import inputCore
import keyboardHandler
from keyboardHandler import KeyboardInputGesture as KIGesture
import NVDAObjects.IAccessible
import speech
import scriptHandler
import textInfos
import time
import tones
from logHandler import log
from nvwave import playWaveFile
# import review
# import screenBitmap
# import ctypes
# import treeInterceptorHandler
from wx import CallLater
import winsound
SCRCAT_AUDACITY = _('Audacity')
MOUSEEVENTF_WHEEL=0x0800
dataPath=os.path.join(os.path.dirname(__file__).decode('mbcs'), 'data\\')
menuFull=[]
assignedShortcuts={}
toolBars={}
lastStatus='Stopped.'
suppressStatus=False

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

	def script_replaceApplications(self, gesture):
		# This is necessary because the usage of the applications key deselects all tracks.
		KIGesture.fromName('shift+f10').send()

	def __init__(self, *args, **kwargs):
		self._audacityInputHelp=False
		super(AppModule, self).__init__(*args, **kwargs)

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
		lookup=assignedShortcuts.get(gesture.displayName, None) 
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
		# Somewhat outdated
		if (windowText=='Track Panel' and windowControlID==1003 and childID>0):
			clsList.insert(0, Track)
			try:
				if name.endswith(' Label Track'):
					clsList.insert(0, labelTrack)
				return
			except:
				KIGesture.fromName('downArrow').send()

	def event_NVDAObject_init(self,obj):
		if obj.windowClassName=='#32770' and obj.role==controlTypes.ROLE_PANE:
			#obj.role=controlTypes.ROLE_DIALOG
			obj.isFocusable=False
		# avoid the ampersand in dialogs
		if obj.windowClassName=='Button' and not obj.role in [controlTypes.ROLE_MENUBAR, controlTypes.ROLE_MENUITEM, controlTypes.ROLE_POPUPMENU]:
			obj.name = api.winUser.getWindowText(obj.windowHandle).replace('&','')
		# define the toolbars as real Toolbars
		# such that the name is spoken automatically on focus entered.
		if obj.role==controlTypes.ROLE_PANE and obj.name and (obj.name.startswith('Audacity ') and obj.firstChild.role!= controlTypes.ROLE_STATUSBAR or obj.name=='Timeline'):
			obj.role=controlTypes.ROLE_TOOLBAR
			obj.name=obj.name.lstrip('Audacity ').rstrip('Toolbar')
		# groupings for controls in e.g. preferences
		if obj and obj.role in [5, 6, 8, 9, 13, 24, 36]:
			# Code snippet by David
			# work around for reading group boxes. In Audacity, any group box is a previous
			# sibling of a control, rather than a previous sibling of the parent of a control.
			# restrict the search two the roles listed above
			groupBox = IAccessibleHandler.findGroupboxObject(obj)
			if groupBox:
				obj.container = groupBox
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

	def _get_Menus(self, obj):
		global menuFull
		if len(menuFull)==0 and obj.previous and obj.previous.role==controlTypes.ROLE_MENUBAR:
			menus=obj.previous.children
			# 9 menus for V2.1.3 and 10(+2) for V2.2.0
			if len(menus) >=9:
				signet=dataPath+'signet.wav'
				playWaveFile(signet)
				for i in range(len(menus)):
					queueHandler.queueFunction(queueHandler.eventQueue,self.getMenuTree,menus[i])

	def _update_Toolbars(self,callback=0):
		global toolBars
		obj=api.getForegroundObject()
		if not obj or obj.name == None:
			try:
				obj=api.getFocusAncestors()[2]
			except:
				obj=api.getFocusObject().parent.parent.parent
			finally:
				return
		if not obj:
			return
		winHandle=obj.windowHandle
		winClass=obj.windowClassName
		if winHandle in toolBars:
			return
		else:
			activeToolBars={tb.name.lstrip('Audacity ').rstrip('Toolbar') : tb for tb in obj.recursiveDescendants if tb.role==controlTypes.ROLE_TOOLBAR}
			if len(activeToolBars)>=13: 
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
		wh=obj.windowHandle
		if wh in toolBars:
			return toolBars[wh].get(key)
		elif obj.windowClassName=='wxWindowNR':
			self._update_Toolbars()
			return toolBars[wh].get(key)

	def event_foreground(self, obj, nextHandler):
		speech.cancelSpeech()
		self._get_Menus(obj)
		queueHandler.queueFunction(queueHandler.eventQueue,self._mapAudacityKeys)
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
		if obj.role in [8, 24] and obj.next and obj.next.role==7 and (obj.next.next==None or not(obj.next.name in obj.next.next.name)):
			try:
				obj.name+=' '+obj.next.name
			except:
				pass
		# suppress things like[Panel, Track View Table, TABLEROW...] FURING FOCUS GAINING
		if obj.windowText =='Track Panel':
			speech.cancelSpeech()
		if obj.windowClassName=='#32768' and obj.role == controlTypes.ROLE_POPUPMENU:
			# focus on first item for context menu
			name, obj.name=obj.name, None
			if name=='DropDown':
				KIGesture.fromName('downArrow').send()
		nextHandler()

	def event_valueChange(self, obj, nextHandler):
		if obj and obj.hasFocus and obj.role==24 and obj.name==None and obj.previous and obj.previous.role==8:
			#if eventHandler.isPendingEvents('valueChange',self):
			tones.beep(3000,20,0,20)
			try:
				# A bug in the Nyquist sliders, there's always a zero appended 
				val=obj.previous.value.rstrip('0')
				val += '0' if val.endswith('.') else ''
				ui.message(val)
			except:
				pass
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
		client.CreateObject("WScript.Shell").Run('sndvol')
	script_systemMixer.__doc__=_('Opens the System Volume Mixer, same as rightclicking on the speaker icon in the tray')
	script_systemMixer.category=SCRCAT_AUDACITY

	def script_states(self,gesture):
		ui.browseableMessage(''.join((x+'\n' for x in toolBars.get(api.getForegroundObject().windowHandle).keys())),'')
		#ui.browseableMessage(repr(toolBars))
	script_states.__doc__=_('Reports all tool bars that are currently docked.')
	script_states.category=SCRCAT_AUDACITY

	def getMenuTree(self, obj):
		global menuFull
		menuFull+=[x for x in obj.recursiveDescendants if x and x.role == 11]

	def script_guide(self, gesture):
		with open(dataPath+'Audacity 2.2.0 Guide.htm','r') as guide:
			guide=guide.read()
			speech.cancelSpeech
			ui.browseableMessage(guide,'Guide',True)
	script_guide.__doc__=_('Shows the famous JAWS Guide for Audacity as browseable document. All quick navigation keys allowed including find dialog and elements list.')
	script_guide.category=SCRCAT_AUDACITY

	def _mapAudacityKeys(self):
		if len(menuFull)>0:
			global assignedShortcuts
			for obj in menuFull:
				#if '\x09' in obj.name:
					cmd=obj.name.rpartition('\x09')
					ncmd=self.replaceMulti(cmd[2], \
						[' ', 'Ctrl', 'Left', 'Right', 'Up', 'Down', 'Pageuparrow', 'Pagedownarrow', 'Return'], \
						['', 'control','leftarrow','rightarrow', 'uparrow', 'downarrow', 'pageup', 'pagedown', 'enter']) 
					ncmd=ncmd.lower()
					try:
						ncmd=KIGesture.fromName(ncmd)
						if cmd[0] in canditatesStartTime or cmd[0] in canditatesEndTime or cmd[0] in canditatesLength:
							self.navGestures[ncmd.identifiers[1]]=self.replaceMulti(cmd[0], [' ','(',')','/'], ['', '', '', ''])
						assignedShortcuts[ncmd.displayName]=(cmd[0], obj, obj.simpleParent.name)
					except KeyError:
						pass

	def script_info(self,gesture):
		#ui.message(str(getPipeFile()['Channels']))
		out=''
		for x in menuFull:
			if x.role==11 and not controlTypes.STATE_HASPOPUP in x.states:
				out+=repr(x.name.partition('\x09')[0])+',\n'
		ui.browseableMessage(out)
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
	# prevent the roles table and row (= trackpanel and audio track) from being spoken 
	controlTypes.silentRolesOnFocus.add(controlTypes.ROLE_TABLEROW)
	controlTypes.silentValuesForRoles.add(controlTypes.ROLE_TABLEROW)
	controlTypes.silentRolesOnFocus.add(controlTypes.ROLE_TABLE)
	shouldAllowIAccessibleFocusEvent=True
	appName='Audacity'

	def __init__(self, *args, **kwargs):
		tones.beep(1000,40)
		super(Track, self).__init__(*args, **kwargs)
		tones.beep(4000,40)
		self.appName=self.appModule.appName

	def initOverlayClass(self):
		#tones.beep(1000,40)
		pass

	def event_gainFocus(self):
		self._update_Toolbars()
		#if self.parent.childCount !=1:
		super(Track,self).event_gainFocus()
		if len(self.navGestures)==0:
			self._mapAudacityKeys
		self.bindGestures(self.navGestures)

	def transportAction(self,button,action=None):
		target=self.transportToolBar(button)
		if not action:
			return target.states
		else:
			try:
				target.doAction(0)
			except NotImplementedError:
				pass
			return


	def menuAction(self,menuHeading,menuItem,action=None):
		target=self.mainMenu(menuHeading).firstChild.getChild(menuItem)
		if not action:
			return target
		else:
			try:
				target.doAction(0)
			except NotImplementedError:
				pass
			return

	def getTransportState(self,button):
		try:
			stateConsts = dict((const, name) for name, const in controlTypes.__dict__.items() if name.startswith('STATE_'))
			ret = ', '.join(
				stateConsts.get(state) or str(state)
				for state in button.states)
		except Exception as e:
			ret = 'exception: %s' % e
		return (ret)

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
			#if controlTypes.STATE_SELECTED in track.states:
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
	def script_SelectiontoStart(self, gesture):
		self.autoTime(gesture,4)
	def script_SelectiontoEnd(self, gesture):
		self.autoTime(gesture,5)
	def script_TrackStarttoCursor(self, gesture):
		self.autoTime(gesture, 4)
	def script_CursortoTrackEnd(self, gesture):
		self.autoTime(gesture, 5)
	#commands that should speak the start or end time of the selection
	def script_AddLabelAtPlaybackPosition(self, gesture):
		self.autoTime(gesture)
	def script_PlayStopandSetCursor(self, gesture):
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
	def script_Shortseekleftduringplayback(self, gesture):
		self.autoTime(gesture)
	def script_Longseekleftduringplayback(self, gesture):
		self.autoTime(gesture)
	def script_SelectionExtendLeft(self, gesture):
		self.autoTime(gesture)
	def script_SetorExtendLeftSelection(self, gesture):
		self.autoTime(gesture,2)
	def script_LeftatPlaybackPosition(self, gesture):
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
	def script_Shortseekrightduringplayback(self, gesture):
		self.autoTime(gesture)
	def script_CursorRight(self, gesture):
		self.autoTime(gesture)
	def script_SelectionExtendRight(self, gesture):
		self.autoTime(gesture,1)
	def script_SetorExtendRightSelection(self, gesture):
		self.autoTime(gesture,3)
	def script_RightatPlaybackPosition(self, gesture):
		self.autoTime(gesture,3)
	def script_SelectionContractRight(self, gesture):
		self.autoTime(gesture,1)
	def script_LongSeekrightduringplayback(self, gesture):
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

class labelTrack(Track):
	shouldAllowIAccessibleFocusEvent=True
	editMode=0
	navMode=0

	def initOverlayClass(self):
		self.isFocusable=True
		#if '><' in self.name:
		# self.editMode=2g

	def event_gainFocus(self):
		super(labelTrack,self).event_gainFocus()