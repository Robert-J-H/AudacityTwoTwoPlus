# Audacity App Module for NVDA
# -*- coding: utf-8 -*-
# Copyright 2017 Robert Hänggi and NVDA Access
import appModuleHandler
from appModules import __path__ as paths
from  appVars import *
import api
import ui
import winUser
import os
import characterProcessing
import controlTypes
import IAccessibleHandler
import eventHandler
import queueHandler
import gui
import inputCore
from keyboardHandler import KeyboardInputGesture as KIGesture
import keyboardHandler
import NVDAObjects.IAccessible
import speech
import scriptHandler
import textInfos
import time
import tones
from logHandler import log
import nvwave
# import review
# import screenBitmap
# import ctypes
# import treeInterceptorHandler
# import wx

SCRCAT_AUDACITY = _('Audacity')
MOUSEEVENTF_WHEEL=0x0800
# better way for obtaining the path?
dataPath=[x for x in paths if 'audacity' in x][0]+'\\data\\'
menuFull=[]
assignedShortcuts={}
toolBars={}
lastStatus=u'Stopped.'
currentInstance=0

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
        data = list(map(lambda x: x.title(), data))
    return dict(zip(data[::2],data[1::2]))

class AppModule(appModuleHandler.AppModule):
    pastTime=currentTime=None
    deltaTime=[]
    tapCounter=0
    tapMedian=200.0
    navGestures={}

    def script_replaceApplications(self, gesture):
        # This is necessary because the usage of the applications key  deselects all tracks.
        KIGesture.fromName('shift+f10').send()

    def __init__(self, *args, **kwargs):
        controlTypes.silentRolesOnFocus.add(controlTypes.ROLE_PANE)
        self._audacityInputHelp=False
        super(AppModule, self).__init__(*args, **kwargs)

    def speakAction(self, action):
        queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, action)

    def event_appModule_gainFocus(self):
        self.bindGesture('kb:applications', 'replaceApplications')
        inputCore.manager._captureFunc = self._inputCaptor

    def event_appModule_loseFocus(self):
        inputCore.manager._captureFunc = None

    def _inputCaptor(self, gesture):
        if api.getFocusObject().windowControlID != 1003 or lastStatus==u'Recording.' or gesture.isNVDAModifierKey or gesture.isModifier:
            return True
        lookup=assignedShortcuts.get(gesture.displayName, None) 
        if lookup and not  lookup[0] in shouldNotAutoSpeak:
            queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, lookup[0])
        return True
        # never comes here
        textList = [gesture.displayName]
        scriptName=u''
        script = gesture.script
        if script:
            scriptName = scriptHandler.getScriptName(script)
            desc = script.__doc__
            if desc:
                textList.append(desc)
        # Punctuation must be spoken for the gesture name (the first chunk) so that punctuation keys are spoken.
        speech.speakText(textList[0], reason=controlTypes.REASON_MESSAGE, symbolLevel=characterProcessing.SYMLVL_ALL)
        for text in textList[1:]:
            speech.speakMessage(text)
        if not self._audacityInputHelp or scriptName == 'toggleAudacityInputHelp':
            return True
        else:
            return False

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
        if (windowText=='Track Panel' and windowControlID==1003 and childID>>0):
            clsList.insert(0, Track)
            try:
                if 'Label' in name:
                    clsList.insert(0, labelTrack)
                return
            except:
                KIGesture.fromName('downArrow').send()

    def playSnippet(self ,fileName):
        nvwave.playWaveFile(dataPath+fileName)

    def event_NVDAObject_init(self,obj):
        #if obj:
        #    eventHandler.requestEvents("nameChange", obj.processID, obj.windowHandle)
        if obj.windowClassName==u'#32770':
            obj.role=controlTypes.ROLE_DIALOG
            obj.isFocusable=True
        if obj.windowClassName=='Button' and not obj.role in [controlTypes.ROLE_MENUBAR, controlTypes.ROLE_MENUITEM, controlTypes.ROLE_POPUPMENU]:
            obj.name = winUser.getWindowText(obj.windowHandle).replace('&','')
        if obj.role==controlTypes.ROLE_PANE and obj.name and (obj.name.startswith('Audacity ') or obj.name=='Timeline'):
            obj.role=controlTypes.ROLE_GROUPING
            obj.name=obj.name.lstrip('Audacity')
        if obj and obj.role in [5, 6, 8, 9, 13, 24, 36]:
            # Code snippet by David
            # work around for reading group boxes. In Audacity, any group box is a previous
            # sibling of a control, rather than a previous sibling of the parent of a control.
            # restrict the search two the roles listed above
            groupBox = IAccessibleHandler.findGroupboxObject(obj)
            if groupBox:
                obj.container = groupBox
        if obj.windowClassName==u'#32768' and obj.role == controlTypes.ROLE_POPUPMENU:
            obj.name='DropDown'
        if obj.role==11 and '\\' in obj.name:
            # rearrange the items in the recent files menu such
            #            that the file name comes before the full qualified path.
            obj.name=u'{2}, {0}{1}{2}'.format(*obj.name.rpartition('\\'))

    def replaceMulti(self, item, old, new):
        while len(old)!=0:
            item=unicode(item).replace(old.pop(0), new.pop(0))
        return item

    def _get_Menus(self, obj):
        global menuFull
        if len(menuFull)==0 and obj.previous and obj.previous.role==controlTypes.ROLE_MENUBAR:
            menus=obj.previous.children
            if len(menus) >=10:
                signet=dataPath+'signet.wav'
                nvwave.playWaveFile(signet)
                for i in range(len(menus)):
                    queueHandler.queueFunction(queueHandler.eventQueue,self.getMenuTree,menus[i])


    def _get_Toolbars(self,obj):
        global toolBars, currentInstance
        winHandle=obj.windowHandle
        if (len(toolBars)==13 and winHandle==currentInstance) or  obj.windowClassName!='wxWindowNR':
            return True
        else:
            currentInstance=winHandle
            try:
                toolBars={tb.name.lstrip(u'Audacity ') : tb for tb in obj.recursiveDescendants  if tb.role==controlTypes.ROLE_GROUPING}
                log.info(str(len(toolBars)))
                if len(toolBars)>=13: 
                    return True
            except:
                pass

    def event_foreground(self, obj, nextHandler):
        speech.cancelSpeech()
        #queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, str(obj.windowClassName))
        self._get_Menus(obj)
        self._get_Toolbars(obj)
        queueHandler.queueFunction(queueHandler.eventQueue,self._mapAudacityKeys)
        nextHandler()

    def event_gainFocus(self, obj, nextHandler):
        # Nyquist effects with an unspoken  unit and an unnamed slider 
        if obj and obj.role==8 and obj.next and obj.next.role==24 and not(obj.next.name) and obj.next.next and obj.previous.location[0]!=obj.next.next.location[0]:
            try:
                obj.name+=' '+obj.next.next.name
            except:
                pass
        if obj and obj.role==24 and not(obj.name) and obj.previous and obj.previous.role==8:
            try:
                obj.name=obj.previous.value
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
        if obj.windowClassName==u'#32768' and obj.role == controlTypes.ROLE_POPUPMENU:
            # focus on first item for context menu
            name, obj.name=obj.name, None
            if name==u'DropDown':
                KIGesture.fromName('downArrow').send()
        nextHandler()

    def event_nameChange(self, obj, nextHandler):
        global lastStatus
        name=obj.name
        log.info(repr(name))
        if name in ['Stopped.','Playing Paused.', 'Recording Paused.'] and name !=lastStatus:
            ui.message(str(name))
        if name in ['Recording.','Playing.','Stopped.','Playing Paused.', 'Recording Paused.']:
            lastStatus=name
        nextHandler()

    def script_states(self,gesture):
        ui.browseableMessage(''.join((x+'\n' for x in toolBars.iterkeys())),'')
        #for elem  in toolBars.viewkeys():ui.message(str(elem))
    script_states.__doc__=_('Reports all tool bars that are currently docked.')
    script_states.category=SCRCAT_AUDACITY

    def getMenuTree(self, obj):
        global menuFull
        menuFull+=[x for x in obj.recursiveDescendants if x and x.role == 11]

    def script_guide(self, gesture):
        with open(str(dataPath+'Audacity 2.1.3 Guide.htm'),'r') as guide:
            guide=guide.read()
            speech.cancelSpeech
            ui.browseableMessage(guide,'Guide',True)
    script_guide.__doc__=_(u'Shows the famous JAWS Guide for Audacity as browseable document. All quick navigation keys allowed including find dialog and elements list.')
    script_guide.category=SCRCAT_AUDACITY

    def _mapAudacityKeys(self):
        if len(menuFull)>0:
            global assignedShortcuts
            for obj  in menuFull:
                #if u'\x09' in obj.name:
                    cmd=obj.name.rpartition(u'\x09')
                    ncmd=self.replaceMulti(cmd[2], \
                        [u' ', u'Ctrl', u'Left', u'Right', u'Up', u'Down', u'Pageuparrow', u'Pagedownarrow', u'Return'], \
                        [u'', u'control',u'leftarrow',u'rightarrow', 'uparrow', u'downarrow', u'pageup', u'pagedown', u'enter']) 
                    ncmd=ncmd.lower()
                    try:
                        ncmd=KIGesture.fromName(ncmd)
                        if cmd[0] in canditatesStartTime or cmd[0] in canditatesEndTime:
                            self.navGestures[ncmd.identifiers[1]]=self.replaceMulti(cmd[0], [u' ',u'(',u')',u'/'], [u'', u'', u'', u''])
                        assignedShortcuts[ncmd.displayName]=(cmd[0],obj)
                    except KeyError:
                        pass

    def script_info(self,gesture):
        #ui.message(str(getPipeFile()['Channels']))
        out=''
        for x in menuFull:
            if x.role==11 and not controlTypes.STATE_HASPOPUP in x.states:
                out+=repr(x.name.partition(u'\x09')[0])+u',\n'
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
        pPeak=toolBars['Playback Meter Toolbar'].firstChild.next.name.partition('Peak')[2] 
        pPeak=toolBars['Playback Meter Toolbar'].firstChild.next.name.partition('Peak')[2] 
        ui.message(pPeak)
    script_announcePlaybackPeak.__doc__=_('Reports the current playback level.')
    script_announcePlaybackPeak.category=SCRCAT_AUDACITY

    def script_announceRecordingPeak(self,gesture):
        repeatCount=scriptHandler.getLastScriptRepeatCount()
        if repeatCount==0:
            rPeak=toolBars['Recording Meter Toolbar'].firstChild.next.name.partition('Peak')[2] 
        ui.message(rPeak)
    script_announceRecordingPeak.__doc__=_('Reports the current recording  level.')
    script_announceRecordingPeak.category=SCRCAT_AUDACITY

    def script_announceTempo(self,gesture):
        repeatCount=scriptHandler.getLastScriptRepeatCount()
        obj=api.getFocusObject()
        if repeatCount==0:
            if obj.role==controlTypes.ROLE_EDITABLETEXT:
                text=unicode(round(60/self.tapMedian,1))
                self._paste_safe(text, obj)
            else:
                ui.message(str(round(60/self.tapMedian,1))+' bpm')
        elif repeatCount==1:
            tempo=(60/self.tapMedian)
            text='Tempo not in classical range'
            for i in range(0, len(self.__tempi)-2,2):
                if tempo>self.__tempi[i] and tempo<=self.__tempi[i+2]:
                    text=self.__tempi[i+1]
            ui.message(text)
    script_announceTempo.__doc__=_('''Reports the  tempo in beats per minute for the last  tapping along.
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
    script_tempoTapping.__doc__=_('use this key  to tap along for some measures. The found tempo can be read out by the announce temp shortcut (normally NVDA+pause).')
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
        winUser.mouse_event(MOUSEEVENTF_WHEEL,0,0,120,None)
    script_wheelForward.__doc__=_('Simulates a turn of the Mouse wheel forward.')
    script_wheelForward.category=SCRCAT_AUDACITY

    def script_wheelBack(self, gesture):
        winUser.mouse_event(MOUSEEVENTF_WHEEL,0,0,-120,None)
    script_wheelBack.__doc__=_('Simulates a turn of the Mouse wheel back.')
    script_wheelBack.category=SCRCAT_AUDACITY

    def script_toggleAudacityInputHelp(self,gesture):
        self._audacityInputHelp=not self._audacityInputHelp
        stateOn = 'Audacity input help on'
        stateOff = 'Audacity input help off'
        state = stateOn if self._audacityInputHelp else stateOff
        ui.message(state)
    script_toggleAudacityInputHelp.__doc__=_('Audacity Turns input help on or off. When on, any input such as pressing a key on the keyboard will tell you what script is associated with that input, if any.')
    script_toggleAudacityInputHelp.category=SCRCAT_AUDACITY

    def getTime(self,child):
        rslt=toolBars.get('Selection Toolbar').getChild(child).name.replace(',','')
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
    #  prevent the roles table and row (= trackpanel and audio track) from being spoken 
    controlTypes.silentRolesOnFocus.add(controlTypes.ROLE_TABLEROW)
    controlTypes.silentRolesOnFocus.add(controlTypes.ROLE_TABLE)
    shouldAllowIAccessibleFocusEvent=True
    appName='Audacity'

    def __init__(self, *args, **kwargs):
        super(Track, self).__init__(*args, **kwargs)
        self.appName=self.appModule.appName
    def initOverlayClass(self):
        pass

    def event_gainFocus(self):
        #if self.parent.childCount !=1:
        super(Track,self).event_gainFocus()
        if len(self.navGestures)==0:
            self._mapAudacityKeys
        self.bindGestures(self.navGestures)
        #inputCore.manager._captureFunc = self._inputCaptor
        self.SelectionToolBar=toolBars.get('Selection Toolbar').getChild
        self.transportToolBar=toolBars.get('Transport Toolbar').getChild
        self.mainMenu=self.parent.parent.parent.previous.getChild

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
            stateConsts = dict((const, name) for name, const in controlTypes.__dict__.iteritems() if name.startswith('STATE_'))
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
        audioPosition=self.getTime(11)
        repeatCount=scriptHandler.getLastScriptRepeatCount()
        if repeatCount==0:
            ui.message(audioPosition)
        elif repeatCount==1:
            speech.speakSpelling(audioPosition)
    script_announceAudioPosition.__doc__=_('Reports the audio  position, both during    Playback and while in stop mode.')
    script_announceAudioPosition.category=SCRCAT_AUDACITY

    def script_announceStart(self, gesture):
        start = self.getTime(14)
        repeatCount=scriptHandler.getLastScriptRepeatCount()
        if repeatCount==0:
            ui.message(start+u' Start')
        elif repeatCount==1:
            speech.speakSpelling(start)
    script_announceStart.__doc__=_('Reports the cursor position or left selection boundary.')
    script_announceStart.category=SCRCAT_AUDACITY

    def script_announceEnd(self, gesture):
        if u'Length' in self.SelectionToolBar(13).value:
            end = self.getTime(15)+' long'
        else:
            end = self.getTime(17)+' end'
        repeatCount=scriptHandler.getLastScriptRepeatCount()
        if repeatCount==0:
            ui.message(end)
        elif repeatCount==1:
            speech.speakSpelling(end)
    script_announceEnd.__doc__=_('Reports the right selection boundary if end is chosen in the selection toolbar. Otherwise, the selection length is reported, followed by the word long.')
    script_announceEnd.category=SCRCAT_AUDACITY
    
    def script_reportSelectedTracks(self,gesture):
        repeatCount=scriptHandler.getLastScriptRepeatCount()
        selChoice=''
        selTracks=[]
        if repeatCount==0:
            selTracks=['Selected are: ']
            selChoice='Select On'
        elif repeatCount==1:
            selTracks=['Muted  are: ']
            selChoice='Mute On'
        elif repeatCount==2:
            selTracks=['Soloed  are: ']
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
            ui.message('No Tracks '+selTracks[0].rstrip(' are: '))
        else:
            ui.message(', '.join(selTracks)) 
    script_reportSelectedTracks.__doc__=_('Reports  the currently selected tracks, if any. Reports muted tracks if pressed twice or soloed tracks if pressed three times.')
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
    script_reduceLeft.__doc__=_('Contracts the selection at the left boundary  and plays a preview.')
    script_reduceLeft.category=SCRCAT_AUDACITY

    def script_expandRight(self, gesture):
        KIGesture.fromName('Shift+RightArrow').send()
        KIGesture.fromName('Shift+F8').send()
    script_expandRight.__doc__=_('Expands the selection at the right boundary  and plays a preview.')
    script_expandRight.category=SCRCAT_AUDACITY

    def script_reduceRight(self, gesture):
        KIGesture.fromName('Control+Shift+LeftArrow').send()
        KIGesture.fromName('Shift+F8').send()
    script_reduceRight.__doc__=_('Contracts the selection at the right boundary  and plays a preview.')
    script_reduceRight.category=SCRCAT_AUDACITY

    def autoTime(self, gesture, which=0):
        gesture.send()
        if lastStatus=='Recording.':
            return
        elif lastStatus=='Playing.':
            if which==2:
                queueHandler.queueFunction(queueHandler.eventQueue,self.playSnippet, 'selection_start.wav')
            elif which==3:
                queueHandler.queueFunction(queueHandler.eventQueue,self.playSnippet, 'selection_end.wav')
        elif lastStatus in [u'Stopped.', u'Playing Paused.', u'blank', None]:
            if which in [0,2]:
                ui.message(self.getTime(11))
            elif which in [1,3]:
                if u'Paused' in lastStatus:
                    ui.message(self.getTime(11))
                else:
                    ui.message(self.getTime(17))
                    pass

    def script_AddLabelAtPlaybackPosition(self, gesture):
        self.autoTime(gesture)
    def script_SelectionStart(self, gesture):
        self.autoTime(gesture)
    def script_SelectionEnd(self, gesture):
        self.autoTime(gesture)
    def script_PlayStopandSetCursor(self, gesture):
        self.autoTime(gesture)
    def script_Pause(self, gesture):
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
        #    self.editMode=2

    def event_gainFocus(self):
        super(labelTrack,self).event_gainFocus()