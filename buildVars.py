# -*- coding: UTF-8 -*-

# Build customizations
# Change this file instead of sconstruct or manifest files, whenever possible.

import time
tl=time.localtime(time.time())

# Full getext (please don't change)
_ = lambda x : x

# Add-on information variables
addon_info = {
	# for previously unpublished addons, please follow the community guidelines at:
	# https://bitbucket.org/nvdaaddonteam/todo/raw/master/guideLines.txt
	# add-on Name, internal for nvda
	"addon_name" : "audacity",
	# Add-on summary, usually the user visible name of the addon.
	# Translators: Summary for this add-on to be shown on installation and add-on information.
	"addon_summary" : _("Audacity 2.2.2"),
	# Add-on description
	# Translators: Long description to be shown for this add-on on add-on information from add-ons manager
	"addon_description" : _("""Provides enhanced Accessibility for the Multi-track editor Audacity, 
such as automatic and manual time announcement, shortcut help and  integration of the JAWS guide."""),
	# version
	"addon_version" : "{}.{}.{}-dev".format("222",str(tl.tm_year)[2:], tl.tm_yday),
	# Author(s)
	"addon_author" : u"Robert Hänggi <aarjay.robert@gmail.com>",
	# URL for the add-on documentation support
	"addon_url" : None,
	# Documentation file name
	"addon_docFileName" : "readme.txt",
}


import os.path

# Define the python files that are the sources of your add-on.
# You can use glob expressions here, they will be expanded.
pythonSources = []

# Files that contain strings for translation. Usually your python sources
i18nSources = pythonSources + ["buildVars.py"]

# Files that will be ignored when building the nvda-addon file
# Paths are relative to the addon directory, not to the root directory of your addon sources.
excludedFiles = []
