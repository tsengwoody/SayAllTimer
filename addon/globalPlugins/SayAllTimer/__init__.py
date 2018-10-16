# coding: utf-8
# SayAllTimer: SayAll stop after certain time
# Copyright (C) 2018 Tseng Woody <tsengwoody.tw@gmail.com>
# This file is covered by the GNU General Public License.
# See the file COPYING.txt for more details.

import api
import config
import controlTypes
import globalPluginHandler
from logHandler import log
import sayAllHandler
import speech
import textInfos
import ui

import time

CURSOR_CARET=0
CURSOR_REVIEW=1

SayAllProfileTrigger = sayAllHandler.SayAllProfileTrigger

def readTextHelper_generator(cursor):
	if cursor==CURSOR_CARET:
		try:
			reader=api.getCaretObject().makeTextInfo(textInfos.POSITION_CARET)
		except (NotImplementedError, RuntimeError):
			return
	else:
		reader=api.getReviewPosition()

	lastSentIndex=0
	lastReceivedIndex=0
	cursorIndexMap={}
	keepReading=True
	speakTextInfoState=speech.SpeakTextInfoState(reader.obj)

	start = time.time()

	with SayAllProfileTrigger():
		while True:
			if not reader.obj:
				# The object died, so we should too.
				return
			# lastReceivedIndex might be None if other speech was interspersed with this say all.
			# In this case, we want to send more text in case this was the last chunk spoken.
			if lastReceivedIndex is None or (lastSentIndex-lastReceivedIndex)<=10:
				if keepReading:
					bookmark=reader.bookmark
					index=lastSentIndex+1
					delta=reader.move(textInfos.UNIT_READINGCHUNK,1,endPoint="end")
					if delta<=0:
						speech.speakWithoutPauses(None)
						keepReading=False
						continue
					speech.speakTextInfo(reader,unit=textInfos.UNIT_READINGCHUNK,reason=controlTypes.REASON_SAYALL,index=index,useCache=speakTextInfoState)
					lastSentIndex=index
					cursorIndexMap[index]=(bookmark,speakTextInfoState.copy())
					try:
						reader.collapse(end=True)
					except RuntimeError: #MS Word when range covers end of document
						# Word specific: without this exception to indicate that further collapsing is not posible, say-all could enter an infinite loop.
						speech.speakWithoutPauses(None)
						keepReading=False
			else:
				# We'll wait for speech to catch up a bit before sending more text.
				if speech.speakWithoutPauses.lastSentIndex is None or (lastSentIndex-speech.speakWithoutPauses.lastSentIndex)>=10:
					# There is a large chunk of pending speech
					# Force speakWithoutPauses to send text to the synth so we can move on.
					speech.speakWithoutPauses(None)
			receivedIndex=speech.getLastSpeechIndex()
			if receivedIndex!=lastReceivedIndex and (lastReceivedIndex!=0 or receivedIndex!=None): 
				lastReceivedIndex=receivedIndex
				bookmark,state=cursorIndexMap.get(receivedIndex,(None,None))
				if state:
					state.updateObj()
				if bookmark is not None:
					updater=reader.obj.makeTextInfo(bookmark)
					if cursor==CURSOR_CARET:
						updater.updateCaret()
					if cursor!=CURSOR_CARET or config.conf["reviewCursor"]["followCaret"]:
						api.setReviewPosition(updater, isCaret=cursor==CURSOR_CARET)
			elif not keepReading and lastReceivedIndex==lastSentIndex:
				# All text has been sent to the synth.
				# Turn the page and start again if the object supports it.
				if isinstance(reader.obj,textInfos.DocumentWithPageTurns):
					try:
						reader.obj.turnPage()
					except RuntimeError:
						break
					else:
						reader=reader.obj.makeTextInfo(textInfos.POSITION_FIRST)
						keepReading=True
				else:
					break

			while speech.isPaused:
				yield
			yield

			now = time.time()

			if (now-start) > int(min)*60 +int(sec):
				speech.cancelSpeech()
				break

		# Wait until the synth has actually finished speaking.
		# Otherwise, if there is a triggered profile with a different synth,
		# we will switch too early and truncate speech (even up to several lines).
		# Send another index and wait for it.
		index=lastSentIndex+1
		speech.speak([speech.IndexCommand(index)])
		while speech.getLastSpeechIndex()<index:
			yield
			yield
		# Some synths say they've handled the index slightly sooner than they actually have,
		# so wait a bit longer.
		for i in xrange(30):
			yield

new_readTextHelper_generator = readTextHelper_generator
old_readTextHelper_generator = sayAllHandler.readTextHelper_generator

def initialize_config():
	config.conf["SayAllTimer"] = {}
	config.conf["SayAllTimer"]["min"] = "30"
	config.conf["SayAllTimer"]["sec"] = "30"

try:
	min = config.conf["SayAllTimer"]["min"]
	sec = config.conf["SayAllTimer"]["sec"]
except:
	initialize_config()
	min = config.conf["SayAllTimer"]["min"]
	sec = config.conf["SayAllTimer"]["sec"]

import gui
from gui import guiHelper
from gui.settingsDialogs import SettingsDialog

import wx

class GeneralSettingsDialog(SettingsDialog):
	# Translators: Title of the SayAllTimerDialog.
	title = _("General Settings")

	def makeSettings(self, settingsSizer):
		sHelper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

		global min
		minLabel = _("&Minute:")
		self.minChoices = [unicode(i) for i in range(60)]
		self.minList = sHelper.addLabeledControl(minLabel, wx.Choice, choices=self.minChoices)
		try:
			index = self.minChoices.index(min)
		except:
			index = 59
		self.minList.Selection = index

		global sec
		secLabel = _("&Second:")
		self.secChoices = [unicode(i) for i in range(60)]
		self.secList = sHelper.addLabeledControl(secLabel, wx.Choice, choices=self.secChoices)
		try:
			index = self.secChoices.index(sec)
		except:
			index = 59
		self.secList.Selection = index

	def postInit(self):
		pass

	def onOk(self,evt):
		global min
		try:
			config.conf["SayAllTimer"]["min"] = min = self.minChoices[self.minList.GetSelection()]
		except:
			config.conf["SayAllTimer"]["min"] = min = 30

		global sec
		try:
			config.conf["SayAllTimer"]["sec"] = sec = self.secChoices[self.secList.GetSelection()]
		except:
			config.conf["SayAllTimer"]["sec"] = sec = 30

		return super(GeneralSettingsDialog, self).onOk(evt)

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = _("SayAllTimer")

	def __init__(self, *args, **kwargs):
		super(GlobalPlugin, self).__init__(*args, **kwargs)
		self.create_menu()

		try:
			self.toggle = config.conf["SayAllTimer"]["toggle"]
		except KeyError:
			config.conf["SayAllTimer"] = {}
			config.conf["SayAllTimer"]["toggle"] = "On"
			self.toggle = config.conf["SayAllTimer"]["toggle"]

		if self.toggle == "On":
			sayAllHandler.readTextHelper_generator = new_readTextHelper_generator
		else:
			sayAllHandler.readTextHelper_generator = old_readTextHelper_generator

	def create_menu(self):
		self.prefsMenu = gui.mainFrame.sysTrayIcon.preferencesMenu
		self.menu = wx.Menu()
		self.generalSettings = self.menu.Append(
			wx.ID_ANY,
			_("&General settings...")
		)
		gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, self.onGeneralSettings, self.generalSettings)

		self.SayAllTimer_item = self.prefsMenu.AppendSubMenu(self.menu, _("SayAllTimer"), _("SayAllTimer"))

	def onGeneralSettings(self, evt):
		gui.mainFrame._popupSettingsDialog(GeneralSettingsDialog)

	def script_toggleSayAllTimer(self,gesture):
		if self.toggle == "On":
			sayAllHandler.readTextHelper_generator = old_readTextHelper_generator
			self.toggle = "Off"
		else:
			sayAllHandler.readTextHelper_generator = new_readTextHelper_generator
			self.toggle = "On"
		ui.message(_("SayAllTimer" +self.toggle))
	script_toggleSayAllTimer.category = scriptCategory
	script_toggleSayAllTimer.__doc__=_("Turns SayAllTimer mode on or off.")

	__gestures={
		#"kb:NVDA+shift+t": "toggleSayAllTimer",
	}
