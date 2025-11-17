#!/usr/bin/env python3

#    parse_top_stats_detailed.py outputs detailed top stats in arcdps logs as parsed by Elite Insights.
#    Copyright (C) 2021 Freya Fleckenstein
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <https://www.gnu.org/licenses/>.


import argparse
import datetime
import os.path
from os import listdir
import sys
import xml.etree.ElementTree as ET
from enum import Enum
import importlib
import xlwt

from collections import OrderedDict
from TW5_parse_top_stats_tools import *

if __name__ == '__main__':
	parser = argparse.ArgumentParser(description='This reads a set of arcdps reports in xml format and generates top stats.')
	parser.add_argument('input_directory', help='Directory containing .xml or .json files from arcdps reports')
	parser.add_argument('-o', '--output', dest="output_filename", help="Text file to write the computed top stats")
	#parser.add_argument('-f', '--input_filetype', dest="filetype", help="filetype of input files. Currently supports json and xml, defaults to json.", default="json")
	parser.add_argument('-x', '--xls_output', dest="xls_output_filename", help="xls file to write the computed top stats")    
	parser.add_argument('-j', '--json_output', dest="json_output_filename", help="json file to write the computed top stats to")    
	parser.add_argument('-l', '--log_file', dest="log_file", help="Logging file with all the output")
	parser.add_argument('-c', '--config_file', dest="config_file", help="Config file with all the settings", default="TW5_parser_config_detailed")
	parser.add_argument('-a', '--anonymized', dest="anonymize", help="Create an anonymized version of the top stats. All account and character names will be replaced.", default=False, action='store_true')
	parser.add_argument('-d', '--extract_date', dest="extract_date", help="Extract date of the logs, different from now")
	args = parser.parse_args()

	if args.extract_date is None:
		myDate = datetime.datetime.now()
	else:
		myDate = datetime.datetime.strptime(args.extract_date, '%Y-%m-%dT%H:%M:%S')


	if not os.path.isdir(args.input_directory):
		print("Directory ",args.input_directory," is not a directory or does not exist!")
		sys.exit()
	if args.output_filename is None:
		args.output_filename = args.input_directory+"/TW5_top_stats_detailed_"+myDate.strftime("%Y%m%d%H%M")+".tid"
	else:
		args.output_filename = args.input_directory+"/"+args.output_filename
	if args.xls_output_filename is None:
		args.xls_output_filename = args.input_directory+"/TW5_top_stats_detailed_"+myDate.strftime("%Y%m%d%H%M")+".xls"
	if args.json_output_filename is None:
		args.json_output_filename = args.input_directory+"/TW5_top_stats_detailed_"+myDate.strftime("%Y%m%d%H%M")+".json"                
	if args.log_file is None:
		args.log_file = args.input_directory+"/log_detailed_"+myDate.strftime("%Y%m%d%H%M")+".txt"

	output = open(args.output_filename, "w",encoding="utf-8")
	log = open(args.log_file, "w")

	parser_config = importlib.import_module("parser_configs."+args.config_file , package=None) 
	
	config = fill_config(parser_config)

	if config.use_PlenBot:
		PlenBotPath = config.PlenBotPath
		getPlenBotLogs(PlenBotPath)
		
	print_string = "Using input directory "+args.input_directory+", writing output to "+args.output_filename+" and log to "+args.log_file
	print(print_string)
	print_string = "Considering fights with at least "+str(config.min_allied_players)+" allied players and at least "+str(config.min_enemy_players)+" enemies that took longer than "+str(config.min_fight_duration)+" s."
	print_to_file(log, print_string)

	players, fights, found_healing, found_barrier, squad_comp, party_comp, squad_offensive, squad_Control, enemy_Control, enemy_Control_Player, downed_Healing, uptime_Table, stacking_uptime_Table, auras_TableIn, auras_TableOut, Death_OnTag, Attendance, DPS_List, CPS_List, SPS_List, HPS_List, DPSStats = collect_stat_data(args, config, log, args.anonymize)    

	# create xls file if it doesn't exist
	book = xlwt.Workbook(encoding="utf-8")
	book.add_sheet("fights overview")
	book.save(args.xls_output_filename)

	
	#Create Tid file header to support drag and drop onto html page
	

	print_to_file(output, 'created: '+myDate.strftime("%Y%m%d%H%M%S"))
	print_to_file(output, 'modified: '+myDate.strftime("%Y%m%d%H%M%S"))
	print_to_file(output, 'creator: '+config.summary_creator)
	print_to_file(output, 'caption: '+myDate.strftime("%Y%m%d%H%M")+'-WvW-Log-Review')
	print_to_file(output, 'tags: Logs [['+myDate.strftime("%Y")+'-'+myDate.strftime("%m")+' Log Reviews]]')
	print_to_file(output, 'title: '+myDate.strftime("%Y%m%d%H%M")+'-WvW-Log-Review\n')
	#End Tid file header

	
	print_string = "__''"+config.summary_title+"''__\n"
	print_to_file(output, print_string)

	# print overall stats
	overall_squad_stats = get_overall_squad_stats(fights, config)
	overall_raid_stats = get_overall_raid_stats(fights)
	total_fight_duration = print_total_squad_stats(fights, overall_squad_stats, overall_raid_stats, found_healing, found_barrier, config, output)

	include_comp_and_review = config.include_comp_and_review
	damage_overview_only = config.damage_overview_only
	defensive_overview_only = config.defensive_overview_only

	DmgOverviewTable = {
        'dmg': "Damage",
        'Pdmg': "Power Dmg",
        'Cdmg': "Condi Dmg",
        'shieldDmg': "Shield Dmg",
        'dmgAll': "Damage All",
		'downContribution': "Down Contrib",
        'againstDownedDamage': "Dmg to Downed",
        'againstDownedCount': "Hits to Downed",
        'downs': "Enemies Downed",
        'kills': "Enemies Killed",
		'dcPct': "Down Contrib PCT of Damage"
    }

	large_items = [
		'<$button setTitle="$:/state/curTab" setTo="Squad Composition" selectedClass="" class="btn btn-sm btn-dark" style=""> Squad Composition </$button>',
		'<$button setTitle="$:/state/curTab" setTo="Fight Review" selectedClass="" class="btn btn-sm btn-dark" style=""> Fight Review </$button>'
	] if include_comp_and_review else []

	#Start nav_bar_menu for TW5
	MenuTabs = ['General', 'Offensive', 'Defensive', 'Support', 'Boons & Buffs', 'Dashboard']

	SubMenuTabs = {
	'General': ['Overview', 'Fight Logs', 'Squad Composition', "Party Composition", 'Fight Review', 'Spike Damage', 'Attendance', 'Support', 'Distance to Tag', 'On Tag Review', 'Skill Casts', 'High Scores', 'Gear Buffs', 'Minions', 'Damage Modifiers', 'Top Skill Dmg'],
	'Offensive': ['Offensive Stats', 'Damage Overview', 'Player Damage by Skill', 'Down Contribution', 'Enemies Downed', 'Enemies Killed', 'Damage', 'Shield Damage', 'Power Damage', 'Condi Damage', 'Against Downed Damage', 'Against Downed Count', 'Damage All', 'DPSStats', 'Burst Damage', 'Damage with Buffs', 'Control Effects - Out', 'Applied Hard CC', 'Applied Hard CC Duration', 'Weapon Swaps'],
	'Defensive': ['Defensive Stats', 'Control Effects - In', 'Condition Uptimes'],
	'Support': ['Healing', 'Barrier', 'Outgoing Healing', 'Condition Cleanses', 'Duration of Conditions Cleansed', 'Stun Breaks', 'Stun Duration Removed', 'Boon Strips', 'Duration of Boons Stripped', 'Illusion of Life', 'Resurrect', 'Downed_Healing', 'Stealth', 'Hide in Shadows', 'FBPages'],
	'Boons & Buffs': ['Total Boons', 'Stability', 'Protection', 'Aegis', 'Might', 'Fury', 'Resistance', 'Resolution', 'Quickness', 'Swiftness', 'Superspeed', 'Alacrity', 'Vigor', 'Regeneration', 'Auras - Out', 'Auras - In', 'Personal Buffs', 'Buff Uptime', 'Stacking Buffs'],
	'Dashboard': ["Dashboard"]
		}

	alertColors = ["primary", "danger", "warning", "success", "info", "light"]

	excludeForMonthly = ['Squad Composition', "Party Composition", 'Fight Review', 'Spike Damage', 'Outgoing Healing', 'Gear Buffs']
	excludeForDmgOverview = ['Down Contribution', 'Enemies Downed', 'Enemies Killed', 'Damage', 'Shield Damage', 'Power Damage', 'Condi Damage', 'Against Downed Damage', 'Against Downed Count', 'Damage All']
	excludeForDefOverview = []

	for item in MenuTabs:
		if not config.charts and item == 'Dashboard':
			continue
		print_to_file(output, '<$button class="btn btn-sm btn-dark"> <$action-setfield $tiddler="$:/state/MenuTab" $field="text" $value="'+item+'"/> <$action-setfield $tiddler="$:/state/curTab" $field="text" $value="'+SubMenuTabs[item][0]+'"/> '+item+' </$button>')
	
	for item in MenuTabs:
		if not config.charts and item == 'Dashboard':
			continue
		print_to_file(output, '<$reveal type="match" state="$:/state/MenuTab" text="'+item+'">')
		print_to_file(output, '\n')
		print_to_file(output, '<<alert-leftbar '+alertColors[MenuTabs.index(item)]+' "'+item+'" width:60%, class:"font-weight-bold">>')
		print_to_file(output, '\n')
		print_to_file(output, '---')
		for tab in SubMenuTabs[item]:
			if not include_comp_and_review and tab in excludeForMonthly:
				continue
			if damage_overview_only and tab in excludeForDmgOverview:
				continue
			if defensive_overview_only and tab in excludeForDefOverview:
				continue
			if not config.charts and item == 'Dashboard':
				continue
			print_to_file(output, '<$button setTitle="$:/state/curTab" setTo="'+tab+'" class="btn btn-sm btn-dark"> '+tab+' </$button>')
		print_to_file(output, '\n')
		print_to_file(output, '</$reveal>')
		print_to_file(output, '\n')			

	#End nav_bar_menu for TW5

	#Overview reveal
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Overview">')
	print_to_file(output, '\n<<alert dark "OVERVIEW" width:60%>>\n')
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')

	print_fights_overview(fights, overall_squad_stats, overall_raid_stats, config, output)

	#End reveal
	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, '</$reveal>')

	#Fight Logs reveal
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Fight Logs">')
	print_to_file(output, '\n<<alert dark "Fight Logs" width:60%>>\n')
	print_to_file(output, '\n---\n')
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')
	print_to_file(output, "|thead-dark table-caption-top sortable|k")
	print_to_file(output, "|Requires Upload to DPSReport to Elite Insight activated to show links|c")
	print_to_file(output, "|Fight# |Date |Start Time |End Time |GMT |Location |Duration |Link |h")
	for fight in Fight_Logs:
		print_string="|"
		for item in fight:
			print_string+=str(item).strip()+" |"
			
		print_to_file(output, print_string)

	#End reveal
	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, '</$reveal>')

	write_fights_overview_xls(fights, overall_squad_stats, overall_raid_stats, config, args.xls_output_filename)

	#Top Damage by Skills for Squad and Enemy
	total_Squad_Damage = sum(total_Squad_Skill_Dmg.values())
	total_Enemy_Damage = sum(total_Enemy_Skill_Dmg.values())	
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Top Skill Dmg">')
	print_to_file(output, '\n<<alert dark "Top Damage by Skill for Squad and Enemy" width:60%>>\n')
	print_to_file(output, "\nCounts based on `player['targetDamageDist']` & `enemy['totalDamageDist']`\n\n")	
	print_to_file(output, '\n---\n')
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')
	print_to_file(output, '<div class="flex-row">')
	print_to_file(output, '    <div class="flex-col border">\n\n')
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|Total Damage by Squad Skill Descending (Top 50)|c")
	print_to_file(output, "|thead-dark table-hover|k")
	print_to_file(output, "|#|!Squad Skill Name | !Damage Output| % of Total|h")
    #start   Squad Skill Damage totals
	sorted_squad_skill_dmg = dict(sorted(total_Squad_Skill_Dmg.items(), key=lambda x: x[1], reverse=True))
	counter = 0
	squadDamageListed = 0
	for name in sorted_squad_skill_dmg:
		counter +=1
		if counter <=50:
			squadDamageListed += sorted_squad_skill_dmg[name]
			pctDmg = round((sorted_squad_skill_dmg[name]/total_Squad_Damage)*100, 2)
			print_to_file(output,'|'+str(counter)+'|'+name+' | '+my_value(sorted_squad_skill_dmg[name])+'| '+my_value(pctDmg)+'%|')
	print_to_file(output, "| Totals|<| "+my_value(squadDamageListed)+"| "+my_value(round((squadDamageListed/total_Squad_Damage)*100,2))+"%|f")
	print_to_file(output, '\n\n\n')
	print_to_file(output, '    </div>')
	print_to_file(output, '<div class="flex-col border">\n\n')
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|Total Damage by Enemy Skill Descending (Top 50)|c")
	print_to_file(output, "|thead-dark table-hover|k")
	print_to_file(output, "|#|!Squad Skill Name | !Damage Output| % of Total|h")
    #start   Enemy Skill Damage totals
	sorted_enemy_skill_dmg = dict(sorted(total_Enemy_Skill_Dmg.items(), key=lambda x: x[1], reverse=True))
	counter = 0
	enemyDamageListed = 0
	for name in sorted_enemy_skill_dmg:
		counter +=1
		if counter <=50:
			enemyDamageListed += sorted_enemy_skill_dmg[name]
			pctDmg = round((sorted_enemy_skill_dmg[name]/total_Enemy_Damage)*100, 2)
			print_to_file(output, '|'+str(counter)+'|'+name+' | '+my_value(sorted_enemy_skill_dmg[name])+'| '+my_value(pctDmg)+'%|')
	print_to_file(output, "| Totals|<| "+my_value(enemyDamageListed)+"| "+my_value(round((enemyDamageListed/total_Enemy_Damage)*100,2))+"%|f")
	print_to_file(output, '\n\n\n')
	print_to_file(output, '    </div>')
	print_to_file(output, '</div>')
	#End reveal - Top Damage by Skills for Squad and Enemy
	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, '</$reveal>')

	#Squad Player Damage by Skills
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Player Damage by Skill">')
	print_to_file(output, '\n<<alert dark "Player Damage by Skill for all Fights" width:60%>>\n')
	print_to_file(output, "\nDamage based on `player['targetDamageDist']`\n\n")	
	print_to_file(output, '\n---\n')
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')
	#print_to_file(output, '<div class="flex-row">')
	#print_to_file(output, '    <div class="flex-col">\n\n')
	#Start Selection Box
	sorted_Player_Damage_by_Skill = OrderedDict(sorted(Player_Damage_by_Skill.items()))

	print_to_file(output, "\n")
	print_to_file(output, "\n")
	print_to_file(output, "<<vspace 25px>>")
	print_to_file(output, "\nSelect Player(s):  ^^Ctrl Click to select multiple^^\n")
	print_to_file(output, "<$select tiddler='$:/state/Player_Selected' default='To View Damage by Skill Table' multiple class='thead-dark'>")
	print_to_file(output, "<option disabled>To View Damage by Skill Table</option>")
	for item in sorted_Player_Damage_by_Skill:
		playerName = sorted_Player_Damage_by_Skill[item]['Name']
		playerProf = sorted_Player_Damage_by_Skill[item]['Prof']
		playerTotal = sorted_Player_Damage_by_Skill[item]['Total']
		if playerTotal < 1:
			continue
		spacedName = playerName.ljust(21, '.')
		print_to_file(output, f'<option style="font-family: monospace">{spacedName}{playerProf}</option>')
	print_to_file(output, "</$select>")
	print_to_file(output, "\n")
	print_to_file(output, "\n<div>")
	print_to_file(output, "\n<$button class='btn btn-sm btn-dark'><$action-setmultiplefields $tiddler='$:/state/Player_Selected' $fields='[[$:/state/Player_Selected]get[text]enlist-input[]]' $values='[[$:/state/Player_Selected]get[text]enlist-input[]]'/>Compare Selected </$button>")
	print_to_file(output, "    ")
	print_to_file(output, "<$button class='btn btn-sm btn-dark'><$action-deletetiddler $tiddler='$:/state/Player_Selected'/>Clear Selected</$button>")
	print_to_file(output, "\n</div>\n")
	print_to_file(output, "---")
	print_to_file(output, "\n")
	print_to_file(output, '\n<div class="flex-row">\n')


	#Start Table Generation
	for item in sorted_Player_Damage_by_Skill:
		playerName = sorted_Player_Damage_by_Skill[item]['Name']
		playerProf = sorted_Player_Damage_by_Skill[item]['Prof']
		playerTotal = sorted_Player_Damage_by_Skill[item]['Total']
		spacedName = playerName.ljust(21, '.')
		if playerTotal < 1:
			continue
		print_to_file(output, f'<$reveal type="match" stateTitle="$:/state/Player_Selected" stateField="{spacedName}{playerProf}" text="{spacedName}{playerProf}">')
		print_to_file(output, '\n<div class="flex-col">\n\n')
		print_to_file(output, "|thead-dark table-hover|k")
		print_to_file(output, "|@@display:block;width:50px;Player@@ | Profession | Total Damage|h")
		print_to_file(output, "|"+playerName+" | {{"+playerProf +"}} | "+my_value(playerTotal)+"|")
		print_to_file(output, "\n\n")
		print_to_file(output, "|thead-dark table-hover|k")
		print_to_file(output, "|@@display:block;width:50px;Skill Name@@ | Damage| % of Total| Min| Avg| Max| Hit| Con Hit| Crit| Crit Dmg| Casts| Hits/Cast|h")
		sorted_Player_Damage_by_Skill_Total = OrderedDict(sorted(Player_Damage_by_Skill[item]['Skills'].items(), key = lambda x: x[1], reverse = True))
		for skill in sorted_Player_Damage_by_Skill_Total:
			skillIcon=""
			for skillID in skill_Dict:
				if skill_Dict[skillID]['name'] == skill:
					skillIcon = skill_Dict[skillID]['icon']
			skillDamage = sorted_Player_Damage_by_Skill_Total[skill][0]
			pctTotal = round((skillDamage / playerTotal)*100,2)
			skill_Min = sorted_Player_Damage_by_Skill_Total[skill][1]
			skill_Max = sorted_Player_Damage_by_Skill_Total[skill][2]
			skill_Hit = sorted_Player_Damage_by_Skill_Total[skill][3]
			skill_connectedHit = sorted_Player_Damage_by_Skill_Total[skill][4]
			skill_Crit = sorted_Player_Damage_by_Skill_Total[skill][5]
			skill_CritDmg = sorted_Player_Damage_by_Skill_Total[skill][6]
			skill_Casts = sorted_Player_Damage_by_Skill_Total[skill][7]
			if skill_connectedHit > 0:
				skill_Avg =int(round((skillDamage / skill_connectedHit),0))
			else:
				skill_Avg = 0
			if skill_connectedHit > 0:
				skill_CritRate = round((skill_Crit / skill_connectedHit)*100,2)
			else:
				skill_CritRate = 0.00
			if skill_Casts >0 and skill_Hit > 0:
				skill_HitsPerCast = round((skill_Hit / skill_Casts),1)
			else:
				skill_HitsPerCast = 0

			print_to_file(output, "|[img width=24 ["+skillIcon+"]] "+skill+" | "+my_value(skillDamage)+"| "+my_value(pctTotal)+"%| "+my_value(skill_Min)+"| "+my_value(skill_Avg)+"| "+my_value(skill_Max)+"| "+my_value(skill_Hit)+"| "+my_value(skill_connectedHit)+"| "+my_value(skill_CritRate)+"%| "+my_value(skill_CritDmg)+"| "+my_value(skill_Casts)+"| "+my_value(skill_HitsPerCast)+"|")
		print_to_file(output, "\n")
		print_to_file(output, "---")
		print_to_file(output, "\n</div>\n")
		print_to_file(output, "\n</$reveal>\n")
	print_to_file(output, '\n\n\n')
	print_to_file(output, '</div>')
	#End reveal - Player Damage by Skills for All Fights
	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, '</$reveal>')


	#Damage Modifier Data Reveal
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Damage Modifiers">')
	print_to_file(output, '\n<<alert dark "Damage Modifiers across all Fights" width:60%>>\n')
	print_to_file(output, "\nCounts based on `player[incomingDamageModifiers]` & `player[damageModifiers]`\n\n")	
	print_to_file(output, '\n---\n')
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')

	tabList = ['Shared', 'Profession']

	DM_Header = ""
	for tab in tabList:
		#make reveal button for each modifier tab in tabList
		DM_Header += '<$button set="$:/state/damModifiers" class="btn btn-sm btn-dark" setTo="'+tab+'">'+tab+' Damage Modifiers</$button> '

	print_to_file(output, DM_Header)
	print_to_file(output, '\n\n---\n\n') 

	#Output Shared Modifier Data
	modListIn = modifierMap['Incoming']['Shared'].keys()
	modListOut = modifierMap['Outgoing']['Shared'].keys()
	print_to_file(output, '\n<$reveal type="match" state="$:/state/damModifiers" text="Shared">\n')
	print_to_file(output, "\n''__Shared Damage Modifiers__''")
	print_to_file(output, "\n---\n\n")    
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|Shared Damage Modifiers |c")
	print_to_file(output, "|thead-dark table-hover sortable|k")
	header = "|!Player Name | !Profession"
	for modifier in modListIn:
		modName = modifier
		modIcon = modifierMap['Incoming']['Shared'][modifier]
		header +=" | ![img width=32 ["+modName+"|"+modIcon+"]]"		
	for modifier in modListOut:
		modName = modifier
		modIcon = modifierMap['Outgoing']['Shared'][modifier]
		header +=" | ![img width=32 ["+modName+"|"+modIcon+"]]"
	header +=" |h"
	print_to_file(output, header)

	for player in squadDamageMods:
		playerName = squadDamageMods[player]['name']
		playerProf = squadDamageMods[player]['profession']
		details = "|"+playerName+" | {{"+playerProf+"}}"
		for modifier in modListIn:
			if modifier in squadDamageMods[player]['Shared']:
				hitCount = squadDamageMods[player]['Shared'][modifier]['hitCount']
				totalHitCount = squadDamageMods[player]['Shared'][modifier]['totalHitCount']
				damageGain = round(squadDamageMods[player]['Shared'][modifier]['damageGain'])
				totalDamage = squadDamageMods[player]['Shared'][modifier]['totalDamage']
				pctHit = my_value(round((hitCount/totalHitCount)*100,2))+"%"
				if totalDamage >0:
					pctDmg = my_value(round(damageGain/(totalDamage+abs(damageGain))*100, 2))+"%"
				else:
					pctDmg = "ToolTip"
				tooltip = str(hitCount)+" out of "+str(totalHitCount)+" hits<br>"+pctHit+" hit<br>Pure Damage: "+my_value(damageGain)
				detailEntry = '<div class="xtooltip">'+pctDmg+' <span class="xtooltiptext">'+tooltip+'</span></div>'
			else:
				detailEntry = "-"
			details += " | "+detailEntry
		for modifier in modListOut:
			if modifier in squadDamageMods[player]['Shared']:
				hitCount = squadDamageMods[player]['Shared'][modifier]['hitCount']
				totalHitCount = squadDamageMods[player]['Shared'][modifier]['totalHitCount']
				damageGain = round(squadDamageMods[player]['Shared'][modifier]['damageGain'])
				totalDamage = squadDamageMods[player]['Shared'][modifier]['totalDamage']
				pctHit = my_value(round((hitCount/totalHitCount)*100,2))+"%"
				if totalDamage >0:
					pctDmg = my_value(round(damageGain/(totalDamage+abs(damageGain))*100, 2))+"%"
				else:
					pctDmg = "ToolTip"
				tooltip = str(hitCount)+" out of "+str(totalHitCount)+" hits<br>"+pctHit+" hit<br>Pure Damage: "+my_value(damageGain)
				detailEntry = '<div class="xtooltip">'+pctDmg+' <span class="xtooltiptext">'+tooltip+'</span></div>'
			else:
				detailEntry = "-"
			details += " | "+detailEntry
		details += " |"
		print_to_file(output, details)
	print_to_file(output, '\n</$reveal>\n')

	#Output Profession Modifier Data
	#modListIn = modifierMap['Incoming']['Prof'].keys()
	#modListOut = modifierMap['Outgoing']['Prof'].keys()
	print_to_file(output, '\n<$reveal type="match" state="$:/state/damModifiers" text="Profession">\n')
	print_to_file(output, "\n''__Profession Damage Modifiers__''")
	print_to_file(output, "\n---\n\n")
	for prof in profModifiers['Professions']:
		print_to_file(output, '<$button setTitle="$:/state/modifierProf" setTo="'+prof+'" selectedClass="" class="btn btn-sm btn-dark" style=""> '+prof+' {{'+prof+'}} </$button>')

	for prof in profModifiers['Professions']:
		print_to_file(output, '\n<$reveal type="match" state="$:/state/modifierProf" text="'+prof+'">\n')
		print_to_file(output, "\n''__"+prof+"__'' {{"+prof+"}}")
		print_to_file(output, "\n---\n\n")		

		modifierList = profModifiers['Professions'][prof]
		header="|Name "

		for modifier in modifierList:
			if modifier in modifierMap['Incoming']['Prof']:
				modName = modifier
				modIcon = modifierMap['Incoming']['Prof'][modifier]
			else:
				modName = modifier
				modIcon = modifierMap['Outgoing']['Prof'][modifier]
			header +=" | ![img width=32 ["+modName+"|"+modIcon+"]]"
		header+=" |h"

		print_to_file(output, header)

		for player in squadDamageMods:
			playerName = squadDamageMods[player]['name']
			playerProf = squadDamageMods[player]['profession']
			playerNameProf = playerName+"{{"+playerProf+"}}"
			if prof in squadDamageMods[player]['profession']:
				details = "|"+playerName
				for modifier in profModifiers['Professions'][prof]:
					if modifier in squadDamageMods[player]['Prof']:
						hitCount = squadDamageMods[player]['Prof'][modifier]['hitCount']
						totalHitCount = squadDamageMods[player]['Prof'][modifier]['totalHitCount']
						damageGain = round(squadDamageMods[player]['Prof'][modifier]['damageGain'])
						totalDamage = squadDamageMods[player]['Prof'][modifier]['totalDamage']
						pctHit = my_value(round((hitCount/totalHitCount)*100,2))+"%"
						if totalDamage >0:
							pctDmg = my_value(round(damageGain/(totalDamage+abs(damageGain))*100, 2))+"%"
						else:
							pctDmg = "ToolTip"
						tooltip = str(hitCount)+" out of "+str(totalHitCount)+" hits<br>"+pctHit+" hit<br>Pure Damage: "+my_value(damageGain)
						detailEntry = '<div class="xtooltip">'+pctDmg+' <span class="xtooltiptext">'+tooltip+'</span></div>'
					else:
						detailEntry = "-"
					details += " | "+detailEntry
				details += " |"
				print_to_file(output, details)


		print_to_file(output, '\n</$reveal>\n')	
		#Start Detail output here
	print_to_file(output, '\n</$reveal>\n')

	#End reveal
	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, '</$reveal>')


	#Minion Data Reveal
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Minions">')
	print_to_file(output, '\n<<alert dark "Player created Minion Data from all Fights" width:60%>>\n')
	print_to_file(output, "\nCounts based on `player[minions][combatReplayData]`\n\n")	
	print_to_file(output, '\n---\n')
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')

	for prof in minion_Data:
		print_to_file(output, '<$button setTitle="$:/state/minionProf" setTo="'+prof+'" selectedClass="" class="btn btn-sm btn-dark" style=""> '+prof+' {{'+prof+'}} </$button>')

	for prof in minion_Data:
		print_to_file(output, '\n<$reveal type="match" state="$:/state/minionProf" text="'+prof+'">\n')
		print_to_file(output, "\n''__"+prof+"__'' {{"+prof+"}}")
		print_to_file(output, "\n---\n\n")    

		minionList = minion_Data[prof]["petsList"]
		header="|Name | !Fights | !Duration "
		for item in minionList:
			header+="| !"+item
		header+=" |h"
		detail=""
		print_to_file(output, "|thead-dark table-caption-top sortable|k")
		print_to_file(output, '| <<hl "Minion Generation" IndianRed>> |c')
		print_to_file(output, header)

		for playerName in minion_Data[prof]["player"]:
			minionDuration = 0
			minionFights = 0
			for nameIndex in players:
				if nameIndex.name == playerName and nameIndex.profession == prof:
					minionDuration = nameIndex.duration_fights_present
					minionFights = nameIndex.num_fights_present

			detail +="|"+playerName+" | "+my_value(minionFights)+" | "+my_value(minionDuration)
			for item in minionList:
				if item in minion_Data[prof]["player"][playerName]:
					detail +=" | "+str(minion_Data[prof]["player"][playerName][item])
				else:
					detail +=" | 0 "
			detail+=" |\n"
		print_to_file(output, detail)
		
		print_to_file(output, '\n</$reveal>\n')
	#End reveal
	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, '</$reveal>')

	#High Scores reveal
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="High Scores">')
	print_to_file(output, '\n<<alert dark "High Scores from all Fights" width:60%>>\n')
	print_to_file(output, "\nStat per second based on `player.stats_per_fight[fight_number]['time_active']`\n\n")	
	print_to_file(output, '\n---\n')
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')
	
	offensiveHighScores = ['dmg_PS', 'downContribution_PS', 'downs_PS', 'kills_PS']
	supportHighScores = ['rips_PS', 'cleanses_PS', 'heal_PS', 'barrier_PS']
	defensiveHighScores = ['dodges_PS', 'evades_PS', 'blocks_PS', 'invulns_PS']
	labelTopFive = {'dmg': 'Damage', 'dmg_PS': 'Damage per Second', 'downContribution': 'Down Contribution', 'downContribution_PS': 'Down Contribution per Second', 'downs': 'Downs', 'downs_PS': 'Downs per Second', 'invulns_PS': 'Invulnerable per Second', 'invulns': 'Invulnerable', 'kills': 'Kills', 'kills_PS': 'Kills per Second', 'rips': 'Boon Strips', 'rips_PS': 'Boon Strips per Second', 'cleanses': 'Condition Cleanses', 'cleanses_PS': 'Condition Cleanses per Second', 'heal': 'Healing', 'heal_PS': 'Healing per Second', 'barrier': 'Barrier', 'barrier_PS': 'Barrier per Second', 'dodges': 'Dodges', 'dodges_PS': 'Dodges per Second', 'evades': 'Evades', 'evades_PS': 'Evades per Second', 'blocks': 'Blocks', 'blocks_PS': 'Blocks per Second', 'downed': 'Downed', 'interupted_PS': 'Downed per Second'}
	#print_to_file(output, '\n\n<<h1 "Offensive High Scores" IndianRed>>\n\n')	
		
	print_to_file(output, '<div class="flex-row">\n\n')

	for stat in offensiveHighScores:
		print_to_file(output, '    <div class="flex-col">\n')
		print_to_file(output, "|thead-dark table-caption-top sortable|k")
		#<<hl "Simple highlight" aqua>>
		print_to_file(output, '| <<hl "'+labelTopFive[stat]+'" IndianRed>> |c')
		print_to_file(output, "|Player |Fight | Score|h")
		if stat in  HighScores:
			sortedHighScore = sorted(HighScores[stat].items(), key = lambda x:x[1], reverse = True)
			for item, value in sortedHighScore:
				print_string="|"
				print_string+=item+" | "
				print_string+=my_value(round(value,2))+"|"
				
				print_to_file(output, print_string)
		print_to_file(output, '\n    </div>')
	print_to_file(output, '\n    </div>')

	#print_to_file(output, '\n\n<<h1 "Support High Scores" LightGreen>>\n\n')
	print_to_file(output, '<div class="flex-row">\n\n')

	for stat in supportHighScores:
		print_to_file(output, '    <div class="flex-col">\n')
		print_to_file(output, "|thead-dark table-caption-top sortable|k")
		print_to_file(output, '| <<hl "'+labelTopFive[stat]+'" LightGreen>> |c')
		print_to_file(output, "|Player |Fight | Score|h")
		if stat in  HighScores:
			sortedHighScore = sorted(HighScores[stat].items(), key = lambda x:x[1], reverse = True)
			for item, value in sortedHighScore:
				print_string="|"
				print_string+=item+" | "
				print_string+=my_value(round(value,2))+"|"
			
				print_to_file(output, print_string)
		print_to_file(output, '\n    </div>')
	print_to_file(output, '\n    </div>')

	#print_to_file(output, '\n\n<<h1 "Defensive High Scores" LightSalmon>>\n\n')
	print_to_file(output, '<div class="flex-row">\n\n')

	for stat in defensiveHighScores:
		print_to_file(output, '    <div class="flex-col">\n')
		print_to_file(output, "|thead-dark table-caption-top sortable|k")
		print_to_file(output, '| <<hl "'+labelTopFive[stat]+'" LightSalmon>> |c')
		print_to_file(output, "|Player |Fight | Score|h")
		if stat in  HighScores:
			sortedHighScore = sorted(HighScores[stat].items(), key = lambda x:x[1], reverse = True)
			for item, value in sortedHighScore:
				print_string="|"
				print_string+=item+" | "
				print_string+=my_value(round(value,2))+"|"
			
				print_to_file(output, print_string)
		print_to_file(output, '\n    </div>')
	print_to_file(output, '\n    </div>')		
	#End reveal
	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, '</$reveal>')
	#End High Scores reveal

	#Gear Buffs reveal
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Gear Buffs">')
	print_to_file(output, '\n<<alert dark "Gear Buff Data from all Fights" width:60%>>\n')
	print_to_file(output, "\nStat per second based on `player.stats_per_fight[fight_number]['time_active']`\n\n")	
	print_to_file(output, '\n---\n')
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')

	print_to_file(output, "!!Gear Buff Uptimes\n\n")
	
	Header = "|thead-dark table-hover table-caption-top sortable|k\n"
	Header += "|Mouseover for details available - Sortable on header click|c\n"
	Header +="|!Player |"
	sortedUsedRelic = OrderedDict(sorted(usedRelicBuff.items()))
	for relicName in sortedUsedRelic:
		headerIcon = ' ![img width=24 ['+relicName+'|'+sortedUsedRelic[relicName]+']] |'
		Header += headerIcon
	Header +="h"
	print_to_file(output, Header)
	details=""
	for player in RelicDataBuffs:
		if RelicDataBuffs[player]:
			details += "|"+player
		else:
			continue
		for relic in sortedUsedRelic:
			if relic in RelicDataBuffs[player] and RelicDataBuffs[player][relic]['buffDuration'] and sum(RelicDataBuffs[player][relic]['fightTime']):
				numFights = "Fights: "+str(len(RelicDataBuffs[player][relic]['fightTime']))
				totalUptime = (sum(RelicDataBuffs[player][relic]['buffDuration']) / sum(RelicDataBuffs[player][relic]['fightTime']))*100
				avgStacks = "Average Stacks: "+str(round(sum(RelicDataBuffs[player][relic]['buffStacks'])/len(RelicDataBuffs[player][relic]['buffStacks']), 3))
				hitData = str(sum(RelicDataBuffs[player][relic]['hitCount']))+" out of "+str(sum(RelicDataBuffs[player][relic]['totalHits']))+" hits"
				damageGain = "Damage Gain: "+my_value(round(sum(RelicDataBuffs[player][relic]['damageGain'])))
				if sum(RelicDataBuffs[player][relic]['buffStacks']):
					tooltip = avgStacks+" <br> "+numFights
				else:
					tooltip = numFights
				if sum(RelicDataBuffs[player][relic]['hitCount']) >0:
					tooltip += " <br> "+hitData
				if sum(RelicDataBuffs[player][relic]['damageGain']) >0:
					tooltip += " <br> "+damageGain
				details += ' | <div class="xtooltip">'+str(round(totalUptime,2))+'% <span class="xtooltiptext">'+tooltip+'</span></div>'
			else:
				details +=' | '
		details +="|\n"
	print_to_file(output, details)
	
	print_to_file(output, "\n\n---\n\n")
	print_to_file(output, "!!Relic Skill Data\n\n")
	
	RelicTableKeys = ['casts', 'totalDamage', 'hits', 'connectedHits', 'crit', 'glance', 'flank', 'missed', 'invulned', 'interrupted', 'evaded', 'blocked', 'shieldDamage', 'critDamage']
	Header = "|thead-dark table-hover table-caption-top sortable|k\n"
	Header += "|Sortable Table, click header to sort|c\n"
	Header +="|!Player |!Relic |"
	for key in RelicTableKeys:
		Header +="!"+ key+" |"
	Header +="h"
	print_to_file(output, Header)

	details=""
	for player in RelicDataSkills:
		for relic in RelicDataSkills[player]:
			details +="|"+player
			details += ' |[img width=24 ['+relic+' |'+usedRelicSkill[relic]+']] - '+relic
			for stat in RelicTableKeys:
				if stat in RelicDataSkills[player][relic]:
					details += " | "+my_value(RelicDataSkills[player][relic][stat])
				else:
					details += " | N/A"
			details+="|\n"
	print_to_file(output, details)
	print_to_file(output, "\n\n---\n\n")
	#End reveal
	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, '</$reveal>')
	#End Gear Buffs Reveal

	#Squad Spike Damage
	if include_comp_and_review:
		print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Spike Damage">\n')    
		print_to_file(output, '\n<<alert dark "SPIKE DAMAGE" width:60%>>\n')
		print_to_file(output, '\n---\n')    
		print_to_file(output, '<div style="overflow-x:auto;">\n\n')

		output_string = "\nSquad Damage output by second (Mouse Scroll to zoom in/out at location)\n"
			
		print_to_file(output, output_string)

		print_to_file(output, '<$echarts $text={{'+myDate.strftime("%Y%m%d%H%M")+'_spike_damage_heatmap_ChartData}} $height="800px" $theme="dark"/>')

		#end reveal
		print_to_file(output, '\n\n</div>\n\n')
		print_to_file(output, "</$reveal>\n")     

	# end Squad Spike Damage

	#Outgoing Healing and Barrier by Target
	if include_comp_and_review:
		print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Outgoing Healing">\n')    
		print_to_file(output, '\n<<alert dark "Outgoing Healing/Barrier by Target" width:60%>>\n')
		print_to_file(output, '\n---\n')    
		print_to_file(output, '<div style="overflow-x:auto;">\n\n')		

		for name in OutgoingHealing:
			print_to_file(output, '<$button setTitle="$:/state/outgoingHealing" setTo="'+name.split("|")[0]+'_'+OutgoingHealing[name]['Prof']+'" selectedClass="" class="btn btn-sm btn-dark" style=""> '+name.split("|")[0]+'{{'+OutgoingHealing[name]['Prof']+'}} </$button>')

		for name in OutgoingHealing:
			totalHealingOutput = 0
			totalBarrierOutput = 0
			for skill in OutgoingHealing[name]['Skills']:
				totalHealingOutput += OutgoingHealing[name]['Skills'][skill][1]
			for skill in OutgoingHealing[name]['Skills_Barrier']:
				totalBarrierOutput += OutgoingHealing[name]['Skills_Barrier'][skill][1]

			healerMaxGroup = max(OutgoingHealing[name]['Group'], key=OutgoingHealing[name]['Group'].get)

			print_to_file(output, '<$reveal type="match" state="$:/state/outgoingHealing" text="'+name.split("|")[0]+'_'+OutgoingHealing[name]['Prof']+'">')
			print_to_file(output, '<div style="overflow-x:auto;">\n\n')
			print_to_file(output, "\n|Healer Name | Party|h")
			print_to_file (output, "|"+name.split("|")[0]+" | "+str(healerMaxGroup)+" |")
			print_to_file(output, "\n\n---\n")
			print_to_file(output, '<div class="flex-row">')
			print_to_file(output, '    <div class="flex-col border">\n')
			print_to_file(output, "|thead-dark table-caption-top sortable|k")
			print_to_file(output, '| <<hl "Total Healing & Barrier by Player" teal>> |c')
			print_to_file(output, "|!Player Name | !Party | !Healing| !Barrier|h")
			for target in OutgoingHealing[name]['Targets']:
				targetMaxGroup = max(OutgoingHealing[name]['Targets'][target]['Group'], key=OutgoingHealing[name]['Targets'][target]['Group'].get)
				if OutgoingHealing[name]['Targets'][target]['Healing'] >0 or OutgoingHealing[name]['Targets'][target]['Barrier']:
					print_to_file(output, "|"+target+" | "+str(targetMaxGroup)+" | "+my_value(OutgoingHealing[name]['Targets'][target]['Healing'])+"| "+my_value(OutgoingHealing[name]['Targets'][target]['Barrier'])+"|")    
			
			print_to_file(output, '\n\n</div>\n\n')
			print_to_file(output, '    <div class="flex-col border">\n')
			print_to_file(output, "|thead-dark table-caption-top sortable|k")
			print_to_file(output, '| <<hl "Total Healing by Skill" lightgreen>> |c')
			print_to_file(output, "|!Skill |!Skill Name | !Hits| !Total Healing| !Heal/Hit| !Pct|h")
			for skill in OutgoingHealing[name]['Skills']:
				hits=OutgoingHealing[name]['Skills'][skill][0]
				heals=OutgoingHealing[name]['Skills'][skill][1]
				skillName = skill_Dict[str(skill)]['name']
				healString = "|"+str(skill)+" |"+str(skillName)+" | "+my_value(hits)+"| "+my_value(heals)+"| "+my_value(round(heals/hits,2))+"| "+my_value(round(heals/totalHealingOutput*100,2))+"%|"
				print_to_file(output, healString)

			print_to_file(output, '\n\n</div>\n\n')
			print_to_file(output, '    <div class="flex-col border">\n')
			print_to_file(output, "|thead-dark table-caption-top sortable|k")
			print_to_file(output, '| <<hl "Total Barrier by Skill" lightblue>> |c')
			print_to_file(output, "|!Skill |!Skill Name | !Hits| !Total Barrier| !Barrier/Hit| !Pct|h")
			for skill in OutgoingHealing[name]['Skills_Barrier']:
				hits=OutgoingHealing[name]['Skills_Barrier'][skill][0]
				heals=OutgoingHealing[name]['Skills_Barrier'][skill][1]
				skillName = skill_Dict[str(skill)]['name']
				healString = "|"+str(skill)+" |"+str(skillName)+" | "+my_value(hits)+"| "+my_value(heals)+"| "+my_value(round(heals/hits,2))+"| "+my_value(round(heals/totalBarrierOutput*100,2))+"%|"
				print_to_file(output, healString)

			print_to_file(output, '\n</div>\n</div>\n</div>\n')
			print_to_file(output, '</$reveal>')

		#end reveal
		print_to_file(output, '\n\n</div>\n\n')
		print_to_file(output, "</$reveal>\n")   	

	#Total Boons
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Total Boons">\n')    
	print_to_file(output, '\n<<alert dark "Total Boon Generation" width:60%>>\n')	
	print_to_file(output, '\n---\n')    
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')
	
	playerCount = len(players)
	calcHeight = str(playerCount*25)
	print_to_file(output, "\n!!Total Boon Generation\n")

	if config.charts:
		print_to_file(output, '<$echarts $text={{'+myDate.strftime("%Y%m%d%H%M")+'_Total_Boon_Generation_BarChartData}} $height="'+calcHeight+'px" $theme="dark"/>')
	else:
		print_to_file(output, '\n Charts Disabled in config \n')

	#end reveal
	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, "</$reveal>\n")     

	# end Total Boons

	#Personal Buffs
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Personal Buffs">\n')    
	print_to_file(output, '\n<<alert dark "Personal Buffs Uptime %" width:60%>>\n')	
	print_to_file(output, '\n---\n')    
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')

	BP_Header = ""
	Prof_String = ""
	Output_String = ""
	print_to_file(output, "|thead-dark|k")
	for profession in buffs_personal:
		BP_Header += '<$button set="$:/state/PersonalBuffs" class="btn btn-sm btn-dark" setTo="'+profession+'">{{'+profession+'}}'+profession+'</$button> '

	print_to_file(output, BP_Header)
	print_to_file(output, '\n\n---\n\n')

	for profession in buffs_personal:
		Prof_Header = "|{{"+profession+"}}Name | !Active Time|"
		for buff in buffs_personal[profession]['buffList']:
			icon = skill_Dict[str(buff)]['icon']
			tooltip = skill_Dict[str(buff)]['name']
			Prof_Header += '![img width=24 tooltip="'+tooltip+'" ['+icon+']]|'
		print_to_file(output, '\n<$reveal type="match" state="$:/state/PersonalBuffs" text="'+profession+'">\n')
		print_to_file(output, "|thead-dark sortable|k")
		print_to_file(output, Prof_Header+"h")
		for playerName in buffs_personal[profession]['player']:
			buffUptimes="|"+playerName+" "
			playerActiveTime = 0
			#get activeTime from players
			for player in players:
				if player.name == playerName and player.profession == profession:
					playerActiveTime = player.duration_active
			buffUptimes+="| "+str(playerActiveTime)
			for buff in buffs_personal[profession]['buffList']:
				if buff in buffs_personal[profession]['player'][playerName].keys() and playerActiveTime>0:
					buffUptimes+="| "+str(round((buffs_personal[profession]['player'][playerName][buff]/playerActiveTime)*100,2))
				else:
					buffUptimes+="| 0.00"
			print_to_file(output, buffUptimes+"|")
		print_to_file(output, "\n</$reveal>\n")

	#end reveal
	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, "</$reveal>\n")     

	# end Personal Bufffs

	#Skill casts
	
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Skill Casts">\n')    
	print_to_file(output, '\n!!!Skill casts / minute\n')
	print_to_file(output, '\n!!!Excludes Auto Attack\n')
	print_to_file(output, '\n---\n')    
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')

	BP_Header = ""
	SkillCast = OrderedDict(sorted(prof_role_skills.items()))
	print_to_file(output, "|thead-dark|k")
	for skillRole in SkillCast:
		BP_Header += '<$button set="$:/state/ProfSkillUsage" class="btn btn-sm btn-dark" setTo="'+skillRole+'">{{'+skillRole.split(' ')[0]+'}}'+skillRole+'</$button> '

	print_to_file(output, BP_Header)
	print_to_file(output, '\n\n---\n\n')

	for skillRole in SkillCast:
		print_to_file(output, '\n<$reveal type="match" state="$:/state/ProfSkillUsage" text="'+skillRole+'">\n')
		print_to_file(output, '\n{{'+skillRole.split(' ')[0]+'}}'+skillRole+'\n')
		skillOrder = sorted(SkillCast[skillRole]['castTotals'].items(), key = lambda x:x[1], reverse = True)
		print_to_file(output, "|thead-dark sortable|k")
		print_string = ("|Name | Fights| ActiveTime| ")
		countSkills = 0

		for key, value in skillOrder:
			if countSkills <30:
				skillIcon = skill_Dict[key]['icon']
				skillName = skill_Dict[key]['name'].replace('"',"'")
				print_string += '![img width=24 tooltip="'+skillName+'" ['+skillIcon+']]|'
				countSkills +=1
		print_string +="h"
		print_to_file(output, print_string)

		for playerName in SkillCast[skillRole]['player']:
			if SkillCast[skillRole]['player'][playerName]['ActiveTime'] == 0:
				continue
			playerFights = SkillCast[skillRole]['player'][playerName]['Fights']
			playerActive = SkillCast[skillRole]['player'][playerName]['ActiveTime']
			print_string = "|"+playerName+" | "+str(playerFights)+"| "+str(playerActive)+"| "
			playerSkills = 0

			for key, value in skillOrder:
				if playerSkills <30:
					if key in SkillCast[skillRole]['player'][playerName]['Skills']:
						totalCasts = SkillCast[skillRole]['player'][playerName]['Skills'][str(key)]
						CastPerMinute = totalCasts/(playerActive/60)
					else:
						CastPerMinute = 0.00
					playerSkills +=1
					print_string += '%.2f' % CastPerMinute+'|'
			print_to_file(output, print_string)
		print_to_file(output, '\n</$reveal>\n')
	
	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, '\n</$reveal>\n')

	# end Skill casts

	if include_comp_and_review:
		#Squad Composition Testing
		print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Squad Composition">')    
		print_to_file(output, '\n<<alert dark "SQUAD COMPOSITION" width:60%>>\n')
		print_to_file(output, '\n<div class="flex-row">\n    <div class="flex-col-2 border">\n\n')
		sort_order = ['Firebrand', 'Scrapper', 'Spellbreaker', "Herald", "Chronomancer", "Reaper", "Scourge", "Dragonhunter", "Guardian", "Elementalist", "Tempest", "Revenant", "Weaver", "Willbender", "Renegade", "Vindicator", "Warrior", "Berserker", "Bladesworn", "Engineer", "Holosmith", "Mechanist", "Ranger", "Druid", "Soulbeast", "Untamed", "Thief", "Daredevil", "Deadeye", "Specter", "Catalyst", "Mesmer", "Mirage", "Virtuoso", "Necromancer", "Harbinger"]

		print_to_file(output, '<div style="overflow-x:auto;">\n\n')

		output_string = ""

		for fight in squad_comp:
			output_string1 = "\n|thead-dark|k\n"
			output_string2 = ""
			output_string1 += "|Fight |"
			output_string2 += "|"+str(fight+1)
			for prof in sort_order:
				if prof in squad_comp[fight]:
					output_string1 += " {{"+str(prof)+"}} |"
					output_string2 += " | "+str(squad_comp[fight][prof])
					
			output_string1 += "h"
			output_string2 += " |\n"
			
			print_to_file(output, output_string1)
			print_to_file(output, output_string2)
		print_to_file(output, '\n\n</div>\n\n')
		print_to_file(output, '\n</div>\n    <div class="flex-col-2 border">\n')
		print_to_file(output, '\n!!!ENEMY COMPOSITION\n')    
		print_to_file(output, '<div style="overflow-x:auto;">\n\n')  
		enemy_squad_num = 0
		for fight in fights:
			if fight.skipped:
				enemy_squad_num += 1
				continue
			enemy_squad_num += 1
			output_string1 = "\n|thead-dark|k\n"
			output_string2 = ""
			output_string1 += "|Fight |"
			output_string2 += "|"+str(enemy_squad_num)
			for prof in sort_order:
				if prof in fight.enemy_squad:
					output_string1 += " {{"+str(prof)+"}} |"
					output_string2 += " | "+str(fight.enemy_squad[prof])

			output_string1 += "h"
			output_string2 += " |\n"

			print_to_file(output, output_string1)
			print_to_file(output, output_string2)
		print_to_file(output, '\n\n</div>\n\n')
		print_to_file(output, '\n</div>\n</div>\n')
		#end reveal
		print_string = "\n</$reveal>\n"
		print_to_file(output, print_string)     


		# end Squad Composition insert

		#Party Composition Testing
		print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Party Composition">')    
		print_to_file(output, '\n<<alert dark "PARTY COMPOSITION by FIGHT" width:60%>>\n')
				
		print_to_file(output, '\n\n')

		output_string = ""

		for fight in party_comp:
			print_to_file(output, "|thead-dark table-hover table-caption-top w-75|k")
			print_to_file(output, "|Fight Number: "+str(fight+1)+" |c")
			print_to_file(output, "| Party | Members |<|<|<|<|h")
			
			#Set details
			details = ""
			for party in party_comp[fight]:
				party_list = party_comp[fight][party]
				chunk_size = 5
				while party_list:
					chunk, party_list = party_list[:chunk_size], party_list[chunk_size:] 
					details +="| "+str(party)
					for i in range(chunk_size):
						if i >=len(chunk):
							details+= " | "
						else:
							details+= " |{{"+str(chunk[i][0])+"}} "+str(chunk[i][1])
					details+=" |\n"
			print_to_file(output, details)
			print_to_file(output, '\n\n')
		#end reveal
		print_to_file(output, "\n</$reveal>\n")     

		# end Party Composition insert

		#start Fight DPS Review insert
		print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Fight Review">')    
		print_to_file(output, '\n<<alert dark "Damage Output Review by Fight-#" width:60%>>\n\n')
		FightNum=0
		for fight in fights:
			FightNum = FightNum+1
			if not fight.skipped:
				print_to_file(output, '<$button setTitle="$:/state/curFight" setTo="Fight-'+str(FightNum)+'" selectedClass="" class="btn btn-sm btn-dark" style=""> Fight-'+str(FightNum)+' </$button>')
		
		print_to_file(output, '\n---\n')
		
		FightNum = 0
		for fight in fights:
			FightNum = FightNum+1
			if not fight.skipped:
				print_to_file(output, '<$reveal type="match" state="$:/state/curFight" text="Fight-'+str(FightNum)+'">')
				print_to_file(output, '\n<div class="flex-row">\n    <div class="flex-col">\n')
				#begin fight summary
				print_to_file(output, "|thead-dark table-hover|k")
				print_to_file(output, "|Fight Summary:| #"+str(FightNum)+"|h")
				print_to_file(output, '|Squad Members: |'+str(fight.squad)+' |')
				print_to_file(output, '|Ally Members: |'+str(fight.notSquad)+' |')				
				print_to_file(output, '|Squad Deaths: |'+str(fight.total_stats['deaths'])+' |')
				print_to_file(output, '|Enemies: |'+str(fight.enemies)+' |')
				print_to_file(output, '|Enemies Downed: |'+str(fight.downs)+' |')
				print_to_file(output, '|Enemies Killed: |'+str(fight.kills)+' |')
				print_to_file(output, '|Fight Duration: |'+str(fight.duration)+' |')
				print_to_file(output, '|Fight End Time: |'+str(fight.end_time)+' |')
				print_to_file(output, '|Squad All Damage: |'+my_value(fight.total_stats['dmgAll'])+' |')
				print_to_file(output, '|Damage Delta (Target/All): |'+my_value(fight.total_stats['dmg'] - fight.total_stats['dmgAll'])+' |')
				print_to_file(output, '|Squad Target Damage: |'+my_value(fight.total_stats['dmg'])+' |')
				if fight.total_stats['dmg']:
					print_to_file(output, '|Squad Shield Damage: |'+my_value(fight.total_stats['shieldDmg'])+'  ('+my_value(round(fight.total_stats['shieldDmg']/fight.total_stats['dmg']*100,1))+'%) |')
				else:
					print_to_file(output, '|Squad Shield Damage: |'+my_value(fight.total_stats['shieldDmg'])+'  ('+my_value(round(fight.total_stats['shieldDmg']/1*100,1))+'%) |')
				print_to_file(output, '|Enemy Target Damage: |'+my_value(fight.total_stats['dmg_taken'])+' |')
				if fight.total_stats['dmg_taken']:
					print_to_file(output, '|Enemy Shield Damage: |'+my_value(fight.total_stats['barrierDamage'])+'  ('+my_value(round(fight.total_stats['barrierDamage']/fight.total_stats['dmg_taken']*100,1))+'%) |')
				else:
					print_to_file(output, '|Enemy Shield Damage: |'+my_value(fight.total_stats['barrierDamage'])+'  ('+my_value(round(fight.total_stats['barrierDamage']/1*100,1))+'%) |')				
				print_to_file(output, '</div>\n\n')
				#Insert Part Composition
				print_to_file(output, '<div class="flex-col-3">\n')
				print_to_file(output, "|thead-dark table-hover table-caption-top w-75|k")
				print_to_file(output, "| Party | Party Members |<|<|<|<|h")

				#Set details
				details = ""
				for party in party_comp[FightNum-1]:
					party_list = party_comp[FightNum-1][party]
					chunk_size = 5
					while party_list:
						chunk, party_list = party_list[:chunk_size], party_list[chunk_size:] 
						details +="| "+str(party)
						for i in range(chunk_size):
							if i >=len(chunk):
								details+= " | "
							else:
								details+= " |{{"+str(chunk[i][0])+"}} "+str(chunk[i][1])
						details+=" |\n"
				print_to_file(output, details)
				print_to_file(output, '\n\n')

				print_to_file(output, '</div>\n</div>\n')
				print_to_file(output, '\n---\n')
				#end fight Summary
				print_to_file(output, '\n<div class="flex-row">\n    <div class="flex-col-1">\n')
				print_to_file(output, "|table-caption-top|k")
				print_to_file(output, "|Damage by Squad Player Descending (Top 20)|c")
				print_to_file(output, "|thead-dark table-hover|k")
				print_to_file(output, "|!Squad Member | !Damage Output|h")
				#begin squad DPS totals
				sorted_squad_Dps = dict(sorted(fight.squad_Dps.items(), key=lambda x: x[1], reverse=True))
				counter = 0
				for name in sorted_squad_Dps:
					counter +=1
					if counter <=20:
						print_to_file(output, '|'+name+'|'+my_value(sorted_squad_Dps[name])+'|')
				#end Squad DPS totals
				print_to_file(output, '\n</div>\n    <div class="flex-col-1">\n')
				print_to_file(output, "|table-caption-top|k")
				print_to_file(output, "|Damage by Squad Skill Descending (Top 20)|c")
				print_to_file(output, "|thead-dark table-hover|k")
				print_to_file(output, "|!Squad Skill Name | !Damage Output|h")
				#start   Squad Skill Damage totals
				sorted_squad_skill_dmg = dict(sorted(fight.squad_skill_dmg.items(), key=lambda x: x[1], reverse=True))
				counter = 0
				for name in sorted_squad_skill_dmg:
					counter +=1
					if counter <=20:
						print_to_file(output, '|'+name+'|'+my_value(sorted_squad_skill_dmg[name])+'|')
				#end Squad Skill Damage totals
				print_to_file(output, '\n</div>\n    <div class="flex-col-1">\n')
				print_to_file(output, "|table-caption-top|k")
				print_to_file(output, "|Damage by Enemy Player Descending (Top 20)|c")            
				print_to_file(output, "|thead-secondary table-hover|k")
				print_to_file(output, "|!Enemy Player | !Damage Output|h")
				#begin Enemy DPS totals
				sorted_enemy_Dps = dict(sorted(fight.enemy_Dps.items(), key=lambda x: x[1], reverse=True))
				counter = 0
				for name in sorted_enemy_Dps:
					counter +=1
					if counter <=20:
						print_to_file(output, '|'+name+'|'+my_value(sorted_enemy_Dps[name])+'|')
				#end Enemy DPS totals
				print_to_file(output, '\n</div>\n    <div class="flex-col-1">\n')
				print_to_file(output, "|table-caption-top|k")
				print_to_file(output, "|Damage by Enemy Skill Descending (Top 20)|c")            
				print_to_file(output, "|thead-secondary table-hover|k")
				print_to_file(output, "|!Enemy Skill | !Damage Output|h")
				#begin Enemy Skill Damage       
				sorted_enemy_skill_dmg = dict(sorted(fight.enemy_skill_dmg.items(), key=lambda x: x[1], reverse=True))
				counter = 0
				for name in sorted_enemy_skill_dmg:
					counter +=1
					if counter <=20:
						print_to_file(output, '|'+name+'|'+my_value(sorted_enemy_skill_dmg[name])+'|')
				#end Enemy Skill Damage
				print_to_file(output, '\n</div>\n</div>\n')
				print_to_file(output, "</$reveal>\n")
		print_to_file(output, "</$reveal>\n")

		#end Fight DPS Review insert

	# print top x players for all stats. If less then x
	# players, print all. If x-th place doubled, print all with the
	# same amount of top x achieved.
	num_used_fights = overall_raid_stats['num_used_fights']

	top_total_stat_players = {key: list() for key in config.stats_to_compute}
	top_consistent_stat_players = {key: list() for key in config.stats_to_compute}
	top_average_stat_players = {key: list() for key in config.stats_to_compute}
	top_percentage_stat_players = {key: list() for key in config.stats_to_compute}
	top_late_players = {key: list() for key in config.stats_to_compute}
	top_jack_of_all_trades_players = {key: list() for key in config.stats_to_compute}    
	
	#JEL-Tweaked to output TW5 formatting (https://drevarr.github.io/FluxCapacity.html)

	for stat in config.stats_to_compute:
		if damage_overview_only and stat in DmgOverviewTable:
			continue
		if defensive_overview_only and tab in config.defenses_to_compute:
			continue
		if stat not in config.aurasOut_to_compute and stat not in config.aurasIn_to_compute and stat not in config.defenses_to_compute:
			if (stat == 'heal' and not found_healing) or (stat == 'barrier' and not found_barrier):
				continue

		fileDate = myDate

		#JEL-Tweaked to output TW5 output to maintain formatted table and slider (https://drevarr.github.io/FluxCapacity.html)
		print_to_file(output,'<$reveal type="match" state="$:/state/curTab" text="'+config.stat_names[stat]+'">')
		if stat in ['dmg', 'Pdmg', 'Cdmg']:
			print_to_file(output, "\n!!!<<alert dark src:'"+config.stat_names[stat].upper()+"  -  Targets Only' width:60%>>\n")
		elif stat == 'dmgAll':
			print_to_file(output, "\n!!!<<alert dark src:'"+config.stat_names[stat].upper()+"  -  includes NPC, Pets, Minions, siege, etc.' width:60%>>\n")
		else:
			print_to_file(output, "\n!!!<<alert dark src:'"+config.stat_names[stat].upper()+"' width:60%>>\n")
		
		if stat == 'dist':
			print_to_file(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
			print_to_file(output, '<div style="overflow-x:auto;">\n\n')
			top_consistent_stat_players[stat] = get_top_players(players, config, stat, StatType.CONSISTENT)
			top_total_stat_players[stat] = get_top_players(players, config, stat, StatType.TOTAL)
			top_average_stat_players[stat] = get_top_players(players, config, stat, StatType.AVERAGE)            
			top_percentage_stat_players[stat],comparison_val = get_and_write_sorted_top_percentage(players, config, num_used_fights, stat, output, StatType.PERCENTAGE, top_consistent_stat_players[stat])
			print_to_file(output, '\n\n\n\n')
			top_percentage_stat_players[stat],comparison_val = get_top_percentage_players(players, config, stat, StatType.PERCENTAGE, num_used_fights, top_consistent_stat_players[stat], top_total_stat_players[stat], list(), list())
			top_average_stat_players[stat] = get_and_write_sorted_average(players, config, num_used_fights, stat, output)			
			print_to_file(output, '\n\n</div>\n\n')
			print_to_file(output, '\n</div>\n    <div class="flex-col border">\n')
			print_to_file(output, '<div style="overflow-x:auto;">\n\n')
			if config.charts:
				print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_'+stat+'_ChartData}} $height="800px" $theme="dark"/>')
			else:
				print_to_file(output, '\n Charts Disabled in config \n')
			print_to_file(output, '\n\n</div>\n\n')
			print_to_file(output, '\n</div>\n</div>\n')
		else:
			print_to_file(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
			print_to_file(output, '<div style="overflow-x:auto;">\n\n')
			if config.player_sorting_stat_type == 'average':
				top_average_stat_players[stat] = get_and_write_sorted_total_by_average(players, config, total_fight_duration, stat, output)
				top_total_stat_players[stat] = get_top_players(players, config, stat, StatType.TOTAL)
			else:
				top_total_stat_players[stat] = get_and_write_sorted_total(players, config, total_fight_duration, stat, output)
				top_average_stat_players[stat] = get_top_players(players, config, stat, StatType.AVERAGE)	
			print_to_file(output, '\n\n\n\n')
			top_consistent_stat_players[stat] = get_and_write_sorted_top_consistent(players, config, num_used_fights, stat, output)			
			print_to_file(output, '\n\n</div>\n\n')
			print_to_file(output, '\n</div>\n    <div class="flex-col border">\n')
			print_to_file(output, '<div style="overflow-x:auto;">\n\n')
			#top_total_stat_players[stat] = get_and_write_sorted_total(players, config, total_fight_duration, stat, output)
			if config.charts:
				print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_'+stat+'_ChartData}} $height="800px" $theme="dark"/>')
			else:
				print_to_file(output, '\n Charts Disabled in config \n')
			print_to_file(output, '\n\n</div>\n\n')
			print_to_file(output, '\n</div>\n</div>\n')
			top_average_stat_players[stat] = get_top_players(players, config, stat, StatType.AVERAGE)
			top_percentage_stat_players[stat],comparison_val = get_top_percentage_players(players, config, stat, StatType.PERCENTAGE, num_used_fights, top_consistent_stat_players[stat], top_total_stat_players[stat], list(), list())
			
		print_to_file(output, "</$reveal>\n")

	#print Auras-Out details
	print_to_file(output,'<$reveal type="match" state="$:/state/curTab" text="Auras - Out">')
	print_to_file(output, '\n!!!<<alert dark src:"Auras - Out" width:60%>>\n')
	for stat in config.aurasOut_to_compute:
		print_to_file(output, '<$button setTitle="$:/state/curAuras-Out" setTo="'+config.stat_names[stat]+'" selectedClass="" class="btn btn-sm btn-dark" style="">'+config.stat_names[stat]+' </$button>')

	for stat in config.aurasOut_to_compute:
		print_to_file(output,'<$reveal type="match" state="$:/state/curAuras-Out" text="'+config.stat_names[stat]+'">')
		print_to_file(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
		print_to_file(output, '<div style="overflow-x:auto;">\n\n')
		if config.player_sorting_stat_type == 'average':
			top_average_stat_players[stat] = get_and_write_sorted_total_by_average(players, config, total_fight_duration, stat, output)
			top_total_stat_players[stat] = get_top_players(players, config, stat, StatType.TOTAL)
		else:
			top_total_stat_players[stat] = get_and_write_sorted_total(players, config, total_fight_duration, stat, output)
			top_average_stat_players[stat] = get_top_players(players, config, stat, StatType.AVERAGE)
		print_to_file(output, '\n\n')
		top_consistent_stat_players[stat] = get_and_write_sorted_top_consistent(players, config, num_used_fights, stat, output)			
		print_to_file(output, '\n</div>')
		print_to_file(output, '\n</div>\n    <div class="flex-col border">\n')
		print_to_file(output, '<div style="overflow-x:auto;">\n')
		if config.charts:
			print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_'+stat+'_ChartData}} $height="800px" $theme="dark"/>')
		else:
			print_to_file(output, '\n Charts Disabled in config \n')	
		print_to_file(output, '\n</div>')
		print_to_file(output, '\n</div></div>\n')
		top_percentage_stat_players[stat],comparison_val = get_top_percentage_players(players, config, stat, StatType.PERCENTAGE, num_used_fights, top_consistent_stat_players[stat], top_total_stat_players[stat], list(), list())
		print_to_file(output, "</$reveal>\n")
	print_to_file(output, "</$reveal>\n")	

	#start Aura Effects Incoming insert
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Auras - In">')    
	print_to_file(output, '\n<<alert-leftbar danger "Auras by receiving Player" width:60%, class:"font-weight-bold">>\n\n')
	Auras_Order = {5677: 'Fire', 5577: 'Shocking', 5579: 'Frost', 5684: 'Magnetic', 25518: 'Light', 39978: 'Dark', 10332: 'Chaos'}
	for Aura in Auras_Order:
		print_to_file(output, '<$button setTitle="$:/state/curAuras-In" setTo="'+Auras_Order[Aura]+'" selectedClass="" class="btn btn-sm btn-dark" style="">'+Auras_Order[Aura]+' Aura </$button>')

	print_to_file(output, '\n---\n')


	for Aura in Auras_Order:
		key = Auras_Order[Aura]
		if key in auras_TableIn:
			sorted_auras_TableIn = dict(sorted(auras_TableIn[key].items(), key=lambda x: x[1], reverse=True))

			i=1

			print_to_file(output, '<$reveal type="match" state="$:/state/curAuras-In" text="'+key+'">\n')
			print_to_file(output, '\n---\n')
			print_to_file(output, "|table-caption-top|k")
			print_to_file(output, "|{{"+key+"}} "+key+" Aura received by Squad Player Descending [TOP 25 Max]|c")
			print_to_file(output, "|thead-dark table-hover sortable|k")
			print_to_file(output, "|!Place |!Name | !Profession | !Total| !Average|h")

			for name in sorted_auras_TableIn:
				prof = "Not Found"
				fightTime = 99999
				counter = 0
				for nameIndex in players:
					if nameIndex.name == name:
						prof = nameIndex.profession
						fightTime = nameIndex.duration_fights_present

				if i <=25:
					print_to_file(output, "| "+str(i)+" |"+name+" | {{"+prof+"}} | "+str(round(sorted_auras_TableIn[name], 4))+"| "+"{:.4f}".format(round(sorted_auras_TableIn[name]/fightTime, 4))+"|")
					i=i+1

			print_to_file(output, "</$reveal>\n")

			write_auras_in_xls(sorted_auras_TableIn, key, players, args.xls_output_filename)
	print_to_file(output, "</$reveal>\n")
	#end Auras Incoming insert

	#print Defense details
	print_to_file(output,'<$reveal type="match" state="$:/state/curTab" text="Defensive Stats">')
	print_to_file(output, '\n!!!<<alert dark src:"Defensive Stats" width:60%>>\n')
	print_to_file(output, '<$button setTitle="$:/state/curDefense" setTo="Overview" selectedClass="" class="btn btn-sm btn-dark" style=""> Defensive Overview </$button>')

	if not defensive_overview_only:
		for stat in config.defenses_to_compute:
			print_to_file(output, '<$button setTitle="$:/state/curDefense" setTo="'+config.stat_names[stat]+'" selectedClass="" class="btn btn-sm btn-dark" style="">'+config.stat_names[stat]+' </$button>')

	#Print Overview Table
	DefensiveOverview = ['dmg_taken', 'barrierDamage', 'hitsMissed', 'interupted', 'invulns', 'evades', 'blocks', 'dodges', 'cleansesIn', 'ripsIn', 'downed', 'deaths', 'receivedCrowdControl','receivedCrowdControlDuration']
	print_to_file(output,'<$reveal type="match" state="$:/state/curDefense" text="Overview">')	
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')
	print_to_file(output, "|thead-dark table-hover sortable|k")	
	print_to_file(output, "|!Name |!Profession | !{{Damage Taken}} | !{{BarrierDamage}} | !Eff {{BarrierDamage}} % | !{{MissedHits}} | !{{Interrupted}} | !{{Invuln}} | !{{Evades}} | !{{Blocks}} | !{{Dodges}} | !{{Condition Cleanses}} | !{{Boon Strips}} | !{{Downed}} | !{{Died}} | !Hard CC| !CC Duration|h")
	for player in players:
		player_name = player.name
		player_prof = player.profession
		print_string = "|"+player_name+"| {{"+player_prof+"}} "
		for item in DefensiveOverview:
			if item == "barrierDamage":
				eff_Damage = player.total_stats["dmg_taken"] or 1
				eff_Barrier = round((player.total_stats[item]/eff_Damage)*100,2)
				print_string += "| "+my_value(player.total_stats[item])+"| "+my_value(eff_Barrier)+"%"
			else:
				print_string += "| "+my_value(player.total_stats[item])
		print_string +="|"
		print_to_file(output, print_string)
	print_to_file(output, '\n</div>')
	print_to_file(output, '\n</$reveal>')
	#overview_only
	for stat in config.defenses_to_compute:
		print_to_file(output,'<$reveal type="match" state="$:/state/curDefense" text="'+config.stat_names[stat]+'">')
		print_to_file(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
		print_to_file(output, '<div style="overflow-x:auto;">\n\n')
		if config.player_sorting_stat_type == 'average':
			top_average_stat_players[stat] = get_and_write_sorted_total_by_average(players, config, total_fight_duration, stat, output)
			top_total_stat_players[stat] = get_top_players(players, config, stat, StatType.TOTAL)
		else:
			top_total_stat_players[stat] = get_and_write_sorted_total(players, config, total_fight_duration, stat, output)
			top_average_stat_players[stat] = get_top_players(players, config, stat, StatType.AVERAGE)
		print_to_file(output, '\n\n')
		top_consistent_stat_players[stat] = get_and_write_sorted_top_consistent(players, config, num_used_fights, stat, output)			
		print_to_file(output, '\n</div>')
		print_to_file(output, '\n</div>\n    <div class="flex-col border">\n')
		print_to_file(output, '<div style="overflow-x:auto;">\n')
		if config.charts:
			print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_'+stat+'_ChartData}} $height="800px" $theme="dark"/>')
		else:
			print_to_file(output, '\n Charts Disabled in config \n')	
		print_to_file(output, '\n</div>')
		print_to_file(output, '\n</div></div>\n')
		top_percentage_stat_players[stat],comparison_val = get_top_percentage_players(players, config, stat, StatType.PERCENTAGE, num_used_fights, top_consistent_stat_players[stat], top_total_stat_players[stat], list(), list())
		print_to_file(output, "</$reveal>\n")
	print_to_file(output, "</$reveal>\n")	
	write_to_json(overall_raid_stats, overall_squad_stats, fights, players, top_total_stat_players, top_average_stat_players, top_consistent_stat_players, top_percentage_stat_players, top_late_players, top_jack_of_all_trades_players, squad_offensive, squad_Control, enemy_Control, enemy_Control_Player, downed_Healing, uptime_Table, stacking_uptime_Table, auras_TableIn, auras_TableOut, Death_OnTag, Attendance, DPS_List, CPS_List, SPS_List, HPS_List, DPSStats, Player_Damage_by_Skill, args.json_output_filename)

	#print table of accounts that fielded support characters
	print_to_file(output,'<$reveal type="match" state="$:/state/curTab" text="Support">')
	print_to_file(output, '\n<<alert dark "Support Players" width:60%>>\n')
	print_to_file(output, "\n")
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')
	# print table header
	print_string = "|thead-dark table-hover sortable|k"    
	print_to_file(output, print_string)
	print_string = "|!Account |!Name |!Profession | !Fights| !Duration|!Support |!Guild Status |h"
	print_to_file(output, print_string)    

	for stat in config.stats_to_compute:
		if (stat == 'rips' or stat == 'cleanses' or stat == 'stability' or stat == 'heal'):
			write_support_players(players, top_total_stat_players[stat], stat, output)

	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, "</$reveal>\n")

	supportCount=0

	#print table of accounts with attendance details
	print_to_file(output,'<$reveal type="match" state="$:/state/curTab" text="Attendance">')
	print_to_file(output, '\n<<alert dark "Attendance" width:60%>>\n')
	print_to_file(output, "\n")
	print_to_file(output, '<div style="overflow-x:auto;">\n\n')
	# print table header
	print_string = "|thead-dark table-hover sortable|k"    
	print_to_file(output, print_string)
	print_string = "|!Account |Prof_Name | Role| !Fights| !Duration| !Guild Status|h"
	print_to_file(output, print_string)    

	for account in Attendance:
		Acct_Fights = Attendance[account]['fights']
		Acct_Duration = Attendance[account]['duration']
		Acct_Guild_Status = Attendance[account]['guildStatus']
		print_string = "|''"+account+"'' | | | ''"+str(Acct_Fights)+"''| ''"+str(Acct_Duration)+"''| ''"+Acct_Guild_Status+"''|h"
		print_to_file(output, print_string)
		for name in Attendance[account]['names']:
			for prof in Attendance[account]['names'][name]['professions']:
				prof_fights = Attendance[account]['names'][name]['professions'][prof]['fights']
				prof_duration = Attendance[account]['names'][name]['professions'][prof]['duration']
				print_string = "| |{{"+prof.split()[0]+"}}"+name+"  | "+prof.split()[1]+" | "+str(prof_fights)+"| "+str(prof_duration)+"| "+Acct_Guild_Status+"|"
				print_to_file(output, print_string)

	print_to_file(output, '\n\n</div>\n\n')
	print_to_file(output, "</$reveal>\n")

	#start Control Effects Outgoing insert
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Control Effects - Out">')    
	print_to_file(output, '\n<<alert dark "Outgoing Control Effects and Conditions generated by the Squad" width:60%>>\n\n')
	#Control_Effects = {720: 'Blinded', 721: 'Crippled', 722: 'Chilled', 727: 'Immobile', 742: 'Weakness', 791: 'Fear', 833: 'Daze', 872: 'Stun', 26766: 'Slow', 27705: 'Taunt', 30778: "Hunter's Mark", 738:"Vulnerability"}
	Control_Effects = {736: 'Bleeding', 737: 'Burning',861: 'Confusion', 723: 'Poison', 19426: 'Torment', 720: 'Blinded', 721: 'Crippled', 722: 'Chilled', 727: 'Immobile', 742: 'Weakness', 791: 'Fear', 833: 'Daze', 872: 'Stun', 26766: 'Slow', 27705: 'Taunt', 738: 'Vulnerability', 30778: "Hunter's Mark", 44633:'Disenchantment', "total1": "appliedCounts", "total2": "totalDuration"}
	for C_E in Control_Effects:
		if C_E == "total2":
			continue
		print_to_file(output, '<$button setTitle="$:/state/curControl-Out" setTo="'+Control_Effects[C_E]+'" selectedClass="" class="btn btn-sm btn-dark" style="">'+Control_Effects[C_E]+' </$button>')
	print_to_file(output, '<$button setTitle="$:/state/curControl-Out" setTo="MOA Tracking" selectedClass="" class="btn btn-sm btn-dark" style="">MOA Tracking </$button>')
	print_to_file(output, '<$button setTitle="$:/state/curControl-Out" setTo="BS_Tracking" selectedClass="" class="btn btn-sm btn-dark" style="">Battle Standard Tracking </$button>')
	print_to_file(output, '\n---\n')
	

	for C_E in Control_Effects:
		if C_E == "total2":
			continue
		key = Control_Effects[C_E]
		if key in squad_Control:
			sorted_squadControl = dict(sorted(squad_Control[key].items(), key=lambda x: x[1], reverse=True))

			i=1
		
			print_to_file(output, '<$reveal type="match" state="$:/state/curControl-Out" text="'+key+'">\n')
			print_to_file(output, '\n---\n')
			print_to_file(output, "|table-caption-top|k")
			print_to_file(output, "|{{"+key+"}} "+key+" output by Squad Player Descending [TOP 25 Max]|c")
			print_to_file(output, "|thead-dark table-hover sortable|k")
			if key == "appliedCounts":
				print_to_file(output, "|!Place |!Name | !Profession | !Total Applied| !Total Duration| !Duration/Applied| !Applied / Sec|h")
			else:
				print_to_file(output, "|!Place |!Name | !Profession | !Total Secs| !Average|h")
			
			for name in sorted_squadControl:
				playerName = name.split("_")[0]
				#prof = name.split("_")[1]
				fightTime = 99999 
				counter = 0
				for nameIndex in players:
					if nameIndex.name+"_{{"+nameIndex.profession+"}}" == name:
						fightTime = nameIndex.duration_fights_present
						prof = nameIndex.profession

				if i <=25:
					if key == "appliedCounts":
						print_to_file(output, "| "+str(i)+" |"+playerName+" | {{"+prof+"}} | "+"{:.4f}".format(round(squad_Control["appliedCounts"][name], 4))+"| "+"{:.4f}".format(round(squad_Control["totalDuration"][name], 4))+"| "+"{:.4f}".format(round((squad_Control["totalDuration"][name]/squad_Control["appliedCounts"][name]), 4))+"| "+"{:.4f}".format(round((squad_Control["appliedCounts"][name]/fightTime), 4))+"|")
					else:
						print_to_file(output, "| "+str(i)+" |"+playerName+" | {{"+prof+"}} | "+"{:.4f}".format(round(sorted_squadControl[name], 4))+"| "+"{:.4f}".format(round((sorted_squadControl[name]/fightTime), 4))+"|")
					i=i+1

			print_to_file(output, "</$reveal>\n")

			write_control_effects_out_xls(sorted_squadControl, key, players, args.xls_output_filename)


	#Add MOA Tracking Tables
	print_to_file(output, '<$reveal type="match" state="$:/state/curControl-Out" text="MOA Tracking">\n')
	print_to_file(output, '\n---\n')
	print_to_file(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
	print_to_file(output, "|thead-dark table-hover sortable table-caption-top|k")
	print_to_file(output, "|MOA Attempts by Squad Player|c")
	print_to_file(output, "|!Name | !Attempted MOA Casting |h")	
	for name in MOA_Casters:
		print_to_file(output, "|"+name+" | "+str(MOA_Casters[name]['attempts'])+" |")
	print_to_file(output, '\n    </div>\n    <div class="flex-col">\n')
	#MOA Target Table
	print_to_file(output, "|thead-dark table-hover sortable table-caption-top|k")
	print_to_file(output, "|Confirmed Missed MOA Attempts by Target|c")
	print_to_file(output, "|!Name | !Missed | !Blocked | !Invulned |h")	
	for name in MOA_Targets:
		print_to_file(output, "|"+name+" | "+str(MOA_Targets[name]['missed'])+" | "+str(MOA_Targets[name]['blocked'])+" | "+str(MOA_Targets[name]['invulned'])+" |")
	print_to_file(output, '\n    </div>\n</div>\n')
	print_to_file(output, "</$reveal>\n")

	#Add Offensive Battle Standard Tracking Tables
	print_to_file(output, '<$reveal type="match" state="$:/state/curControl-Out" text="BS_Tracking">\n')
	print_to_file(output, '\n---\n')
	print_to_file(output, '\n<div class="flex-row">\n    <div class="flex-col">\n')
	print_to_file(output, "|thead-dark table-hover sortable table-caption-top|k")
	print_to_file(output, "|Offensive Battle Standard Attempts by Squad Player|c")
	print_to_file(output, "|!Player|!Damage|!Hits|!Connected Hits|!Crit|!Missed|!Invulned|!Interrupted|!Evaded|!Blocked|!Shield Damage|!Crit Damage|h")	
	for name in battle_Standard:
		print_to_file(output, "|"+name+" | "+my_value(battle_Standard[name]['totalDamage'])+"| "+my_value(battle_Standard[name]['hits'])+"| "+my_value(battle_Standard[name]['connectedHits'])+"| "+my_value(battle_Standard[name]['crit'])+"| "+my_value(battle_Standard[name]['missed'])+"| "+my_value(battle_Standard[name]['invulned'])+"| "+my_value(battle_Standard[name]['interrupted'])+"| "+my_value(battle_Standard[name]['evaded'])+" |"+my_value(battle_Standard[name]['blocked'])+"| "+my_value(battle_Standard[name]['shieldDamage'])+"| "+my_value(battle_Standard[name]['critDamage'])+"|")
	print_to_file(output, '\n    </div>\n')
	print_to_file(output, '\n    </div>\n')
	print_to_file(output, "</$reveal>\n")

	print_to_file(output, "</$reveal>\n")	
	#end Control Effects Outgoing insert

	#start Control Effects Incoming insert
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Control Effects - In">')    
	print_to_file(output, '\n<<alert dark "Incoming Control Effects generated by the Enemy" width:60%>>\n\n')
	Control_Effects = {720: 'Blinded', 721: 'Crippled', 722: 'Chilled', 727: 'Immobile', 742: 'Weakness', 791: 'Fear', 833: 'Daze', 872: 'Stun', 26766: 'Slow', 27705: 'Taunt', 30778: "Hunter's Mark", 738:"Vulnerability", 44633:'Disenchantment'}
	for C_E in Control_Effects:
		print_to_file(output, '<$button setTitle="$:/state/curControl-In" setTo="'+Control_Effects[C_E]+'" selectedClass="" class="btn btn-sm btn-dark" style="">'+Control_Effects[C_E]+' </$button>')
	
	print_to_file(output, '\n---\n')
	

	for C_E in Control_Effects:
		key = Control_Effects[C_E]
		if key in enemy_Control:
			sorted_enemyControl = dict(sorted(enemy_Control[key].items(), key=lambda x: x[1], reverse=True))

			i=1
			
			print_to_file(output, '<$reveal type="match" state="$:/state/curControl-In" text="'+key+'">\n')
			print_to_file(output, '\n---\n')
			print_to_file(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
			print_to_file(output, "|table-caption-top|k")
			print_to_file(output, "|{{"+key+"}} "+key+" impacted Squad Player Descending [TOP 25 Max]|c")
			print_to_file(output, "|thead-dark table-hover sortable|k")
			print_to_file(output, "|!Place |!Name | !Profession | !Total| !Average|h")
			
			for name in sorted_enemyControl:
				prof = "Not Found"
				fightTime = 99999 
				counter = 0
				for nameIndex in players:
					if nameIndex.name == name:
						prof = nameIndex.profession
						fightTime = nameIndex.duration_fights_present

				if i <=25:
					print_to_file(output, "| "+str(i)+" |"+name+" | {{"+prof+"}} | "+str(round(sorted_enemyControl[name], 4))+"| "+"{:.4f}".format(round(sorted_enemyControl[name]/fightTime, 4))+"|")
					i=i+1

			#print_to_file(output, "</$reveal>\n")

			write_control_effects_in_xls(sorted_enemyControl, key, players, args.xls_output_filename)

		if key in enemy_Control_Player:
			sorted_enemyControlPlayer = dict(sorted(enemy_Control_Player[key].items(), key=lambda x: x[1], reverse=True))

			i=1
	
			print_to_file(output, '\n---\n')
			print_to_file(output, '\n</div>\n    <div class="flex-col border">\n')
			print_to_file(output, "|table-caption-top|k")
			print_to_file(output, "|{{"+key+"}} "+key+" output by Enemy Player Descending [TOP 25 Max]|c")
			print_to_file(output, "|thead-dark table-hover sortable|k")
			print_to_file(output, "|!Place |!Name | !Profession | !Total|h")
		
			for name in sorted_enemyControlPlayer:
				prof = name.split(' pl')[0]
				counter = 0

				if i <=25:
					print_to_file(output, "| "+str(i)+" |"+name+" | {{"+prof+"}} | "+str(round(sorted_enemyControlPlayer[name],4 ))+"|")
					i=i+1

			print_to_file(output, '\n</div>\n</div>\n')
			print_to_file(output, "</$reveal>\n")

	print_to_file(output, "</$reveal>\n")
	#end Control Effects Incoming insert

	#start Condition Uptimes insert
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Condition Uptimes">')    
	print_to_file(output, '\n<<alert dark "Condition Uptimes generated by the Enemy" width:60%>>\n\n')
	condition_ids = {736: 'Bleeding', 737: 'Burning',861: 'Confusion', 723: 'Poison', 19426: 'Torment', 720: 'Blinded', 721: 'Crippled', 722: 'Chilled', 727: 'Immobile', 742: 'Weakness', 791: 'Fear', 833: 'Daze', 872: 'Stun', 26766: 'Slow', 27705: 'Taunt', 738: 'Vulnerability'}
	print_to_file(output, '\n---\n')
	print_to_file(output, '<<tc src:"Uptimes in Green reflect further reduction due to Resistance, hover for the original uptime." color:"green">>')

	TableKeys = ['Bleeding', 'Burning', 'Confusion', 'Poison', 'Torment', 'Blinded', 'Crippled', 'Chilled', 'Immobile', 'Weakness', 'Fear', 'Daze', 'Stun', 'Slow', 'Taunt', 'Vulnerability']
	Header = "|thead-dark table-hover table-caption-top sortable|k\n"
	Header += "|Squad Composite Uptime % for Conditions |c\n"
	Header +="|!@@display:block;width:137px;Squad Data@@ |!@@display:block;width:137px;FightTime@@ | !@@display:block;width:137px;Incoming Clears@@ |"
	for key in TableKeys:
		Header += " !@@display:block;width:40px;{{"+key+"}}@@ |"
	Header +="h"
	print_to_file(output, Header)
	
	details="|Squad Totals:| "+str(round(conditionDataSquad['totalFightTime'],1))+" | "+str(round(conditionDataSquad['IncomingClears']))
	for stat in TableKeys:
		if stat in conditionDataSquad:
			if stat in ResistanceData['Squad'] and ResistanceData['Squad'][stat]:
				if round((((conditionDataSquad[stat]-ResistanceData['Squad'][stat])/conditionDataSquad['totalFightTime'])*100),1) <=0:
					resistReduced = "0.0"
				else:
					resistReduced = str(round((((conditionDataSquad[stat]-ResistanceData['Squad'][stat])/conditionDataSquad['totalFightTime'])*100),1))
				details += ' | <span class="tooltip" data-tooltip=" Uptime w/o Resist: '+str(round(((conditionDataSquad[stat]/conditionDataSquad['totalFightTime'])*100),1))+'%"> @@color:green;'+resistReduced+'%@@</span>'
			else:
				details += " | "+str(round(((conditionDataSquad[stat]/conditionDataSquad['totalFightTime'])*100),1))+"%"
		else:
			details += " | 0.0%"
	details+=" |\n"
	print_to_file(output, details)
	print_to_file(output, '\n---\n')
	
	Header = "|thead-dark table-hover table-caption-top sortable|k\n"
	Header += "|Party Composite Uptime % for Conditions |c\n"
	Header +="|!@@display:block;width:137px;Group Data@@ | !@@display:block;width:137px;FightTime@@ | !@@display:block;width:137px;Incoming Clears@@ |"
	for key in TableKeys:
		Header += " !@@display:block;width:40px;{{"+key+"}}@@ |"
	Header +="h"
	print_to_file(output, Header)
	
	details=""
	for player in conditionDataGroups:
		details += "|Group "+str(player)+" | "+str(round(conditionDataGroups[player]['totalFightTime'],1))+" | "+str(round(conditionDataGroups[player]['IncomingClears']))
		for stat in TableKeys:
			if stat in conditionDataGroups[player]:
				if stat in ResistanceData['Group'][player]['ResistOffset'] and ResistanceData['Group'][player]['ResistOffset'][stat]:
					if round((((conditionDataGroups[player][stat]-ResistanceData['Group'][player]['ResistOffset'][stat])/conditionDataGroups[player]['totalFightTime'])*100),1) <=0:
						resistReduced = "0.0"
					else:
						resistReduced = str(round((((conditionDataGroups[player][stat]-ResistanceData['Group'][player]['ResistOffset'][stat])/conditionDataGroups[player]['totalFightTime'])*100),1))
					details += ' | <span class="tooltip" data-tooltip=" Uptime w/o Resist: '+str(round(((conditionDataGroups[player][stat]/conditionDataGroups[player]['totalFightTime'])*100),1))+'%"> @@color:green;'+resistReduced+'%@@</span>'
				else:
					details += " | "+str(round(((conditionDataGroups[player][stat]/conditionDataGroups[player]['totalFightTime'])*100),1))+"%"
			else:
				details += " | 0.0%"
		details+=" |\n"
	print_to_file(output, details)
	print_to_file(output, '\n---\n')
	
	Header = "|thead-dark table-hover table-caption-top sortable|k\n"
	Header += "|Player Composite Uptime % for Conditions |c\n"
	Header +="|!@@display:block;width:60px;Prof@@ |!@@display:block;width:176px;Player@@ |!@@display:block;width:60px;Group@@ |!@@display:block;width:100px;FightTime@@ |"
	for key in TableKeys:
		Header += " !@@display:block;width:40px;{{"+key+"}}@@ |"
	Header +="h"
	print_to_file(output, Header)
	
	details=""
	for player in conditionData:
		details +=player+" | "+str(round(conditionData[player]['totalFightTime'],1))
		for stat in TableKeys:
			if stat in conditionData[player]:
				if stat in ResistanceData[player]['ResistOffset'] and ResistanceData[player]['ResistOffset'][stat]:
					if round((((conditionData[player][stat]-ResistanceData[player]['ResistOffset'][stat])/conditionData[player]['totalFightTime'])*100),1) <=0:
						resistReduced = "0.0"
					else:
						resistReduced = str(round((((conditionData[player][stat]-ResistanceData[player]['ResistOffset'][stat])/conditionData[player]['totalFightTime'])*100),1))
					details += ' | <span class="tooltip" data-tooltip=" Uptime w/o Resist: '+str(round(((conditionData[player][stat]/conditionData[player]['totalFightTime'])*100),1))+'%"> @@color:green;'+resistReduced+'%@@</span>'
				else:
					details += " | "+str(round(((conditionData[player][stat]/conditionData[player]['totalFightTime'])*100),1))+"%"
			else:
				details += " | 0.0%"
		details+=" |\n"
	print_to_file(output, details)
	print_to_file(output, '\n---\n')
			
	print_to_file(output, "</$reveal>\n")
	#end Condition Uptimes insert	

	#start Buff Uptime Table insert
	uptime_Order = ['stability',  'protection',  'aegis',  'might',  'fury',  'resistance',  'resolution',  'quickness',  'swiftness',  'alacrity',  'vigor',  'regeneration']
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Buff Uptime">')    
	print_to_file(output, '\n<<alert dark "Total Buff Uptime % across all fights" width:60%>>\n\n')
	
	print_to_file(output, '\n---\n')
	print_to_file(output, '<$button setTitle="$:/state/curUptime" setTo="Player" selectedClass="" class="btn btn-sm btn-dark" style=""> Player Uptime Overview </$button>')
	print_to_file(output, '<$button setTitle="$:/state/curUptime" setTo="Party" selectedClass="" class="btn btn-sm btn-dark" style=""> Squad/Party Uptime Overview </$button>')
	print_to_file(output, '\n---\n')

	print_to_file(output, '<$reveal type="match" state="$:/state/curUptime" text="Player">\n')
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|Sortable table - Click header item to sort table |c")
	print_to_file(output, "|thead-dark table-hover sortable|k")
	print_to_file(output, "|!Name | !Profession | !Attendance| !{{Stability}}|  !{{Protection}}|  !{{Aegis}}|  !{{Might}}|  !{{Fury}}|  !{{Resistance}}|  !{{Resolution}}|  !{{Quickness}}|  !{{Swiftness}}|  !{{Alacrity}}|  !{{Vigor}}|  !{{Regeneration}}|h")
	for squadDps_prof_name in uptime_Table:
		fightTime = uptime_Table[squadDps_prof_name]['duration']
		name = uptime_Table[squadDps_prof_name]['name']
		prof = uptime_Table[squadDps_prof_name]['prof']

		output_string = "|"+name+" |"
		output_string += " {{"+prof+"}} | "+my_value(round(fightTime))+"|"
		for item in uptime_Order:
			if item in uptime_Table[squadDps_prof_name] and fightTime >0:
				output_string += " "+"{:.2f}".format(round(((uptime_Table[squadDps_prof_name][item]/fightTime)*100), 2))+"%|"
			else:
				output_string += " 0.00%|"
				


		print_to_file(output, output_string)
	print_to_file(output, "\n</$reveal>\n")

	#Display Squad and Party Buff Uptime Tables
	print_to_file(output, '<$reveal type="match" state="$:/state/curUptime" text="Party">\n')
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|Squad Composite Uptime % for Buffs |c")
	print_to_file(output, "|thead-dark table-hover sortable|k")
	print_to_file(output, "|!Squad/Party | !{{Stability}}|  !{{Protection}}|  !{{Aegis}}|  !{{Might}}|  !{{Fury}}|  !{{Resistance}}|  !{{Resolution}}|  !{{Quickness}}|  !{{Swiftness}}|  !{{Alacrity}}|  !{{Vigor}}|  !{{Regeneration}}|h")
	output_string = "|Squad |"
	for item in uptime_Order:
		if item in squadUptimes['buffs']:
			output_string += " "+"{:.2f}".format(round(((squadUptimes['buffs'][item]/squadUptimes['FightTime'])*100), 2))+"%|"
		else:
			output_string += " 0.00%|"
	print_to_file(output, output_string)
	#print_to_file(output, "</$reveal>\n")

	#print_to_file(output, "\n</$reveal>\n")
	print_to_file(output, '\n---\n')
	#party Uptimes
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|Party Composite Uptime % for Buffs - Sortable table - Click header item to sort table |c")
	print_to_file(output, "|thead-dark table-hover sortable|k")
	print_to_file(output, "|!Squad/Party | !{{Stability}}|  !{{Protection}}|  !{{Aegis}}|  !{{Might}}|  !{{Fury}}|  !{{Resistance}}|  !{{Resolution}}|  !{{Quickness}}|  !{{Swiftness}}|  !{{Alacrity}}|  !{{Vigor}}|  !{{Regeneration}}|h")
	for party in partyUptimes:
		output_string = "|Party: "+str(party)+" |"

		for item in uptime_Order:
			if item in partyUptimes[party]['buffs']:
				if partyUptimes[party]['totalFightTime']:
					output_string += " "+"{:.2f}".format(round(((partyUptimes[party]['buffs'][item]/partyUptimes[party]['totalFightTime'])*100), 2))+"%|"
				else:
					output_string += " 0.00%|"
			else:
				output_string += " 0.00%|"
		print_to_file(output, output_string)
	print_to_file(output, "</$reveal>\n")

	write_buff_uptimes_in_xls(uptime_Table, players, uptime_Order, args.xls_output_filename)
	print_to_file(output, "\n</$reveal>\n")

	#end Buff Uptime Table insert

	max_fightTime = 0
	for squadDps_prof_name in DPSStats:
		max_fightTime = max(DPSStats[squadDps_prof_name]['duration'], max_fightTime)
	
	#start Stacking Buff Uptime Table insert
	stacking_buff_Order = ['might', 'stability']
	max_stacking_buff_fight_time = 0
	for uptime_prof_name in stacking_uptime_Table:
		max_stacking_buff_fight_time = max(stacking_uptime_Table[uptime_prof_name]['duration_might'], max_stacking_buff_fight_time)
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Stacking Buffs">')    
	print_to_file(output, '\n<<alert dark "Stacking Buffs" width:60%>>\n\n')
	for stacking_buff in stacking_buff_Order:
		print_to_file(output, '<$button setTitle="$:/state/curStackingBuffs" setTo="'+stacking_buff+'" selectedClass="" class="btn btn-sm btn-dark" style="">'+stacking_buff+'</$button>')
	
	print_to_file(output, '\n---\n')

	# Might stack table
	print_to_file(output, '<$reveal type="match" state="$:/state/curStackingBuffs" text="might">\n')
	print_to_file(output, '\n---\n')
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|{{Might}} uptime by stack|c")
	print_to_file(output, '|thead-dark table-hover sortable|k')
	output_header =  '|!Name | !Class'
	output_header += ' | ! <span data-tooltip="Number of seconds player was in squad logs">Seconds</span>'
	output_header += '| !Avg| !1+ %| !5+ %| !10+ %| !15+ %| !20+ %| !25 %'
	output_header += '|h'
	print_to_file(output, output_header)
	
	might_sorted_stacking_uptime_Table = []
	for uptime_prof_name in stacking_uptime_Table:
		fight_time = (stacking_uptime_Table[uptime_prof_name]['duration_might'] / 1000) or 1
		might_stacks = stacking_uptime_Table[uptime_prof_name]['might']

		if (DPSStats[uptime_prof_name]['duration'] * 100) / max_fightTime < config.min_attendance_percentage_for_top:
			continue

		avg_might = sum(stack_num * might_stacks[stack_num] for stack_num in range(1, 26)) / (fight_time * 1000)
		might_sorted_stacking_uptime_Table.append([uptime_prof_name, avg_might])
	might_sorted_stacking_uptime_Table = sorted(might_sorted_stacking_uptime_Table, key=lambda x: x[1], reverse=True)
	might_sorted_stacking_uptime_Table = list(map(lambda x: x[0], might_sorted_stacking_uptime_Table))
	
	for uptime_prof_name in might_sorted_stacking_uptime_Table:
		name = stacking_uptime_Table[uptime_prof_name]['name']
		prof = stacking_uptime_Table[uptime_prof_name]['profession']
		fight_time = (stacking_uptime_Table[uptime_prof_name]['duration_might'] / 1000) or 1
		might_stacks = stacking_uptime_Table[uptime_prof_name]['might']

		avg_might = sum(stack_num * might_stacks[stack_num] for stack_num in range(1, 26)) / (fight_time * 1000)
		might_uptime = 1.0 - (might_stacks[0] / (fight_time * 1000))
		might_5_uptime = sum(might_stacks[i] for i in range(5,26)) / (fight_time * 1000)
		might_10_uptime = sum(might_stacks[i] for i in range(10,26)) / (fight_time * 1000)
		might_15_uptime = sum(might_stacks[i] for i in range(15,26)) / (fight_time * 1000)
		might_20_uptime = sum(might_stacks[i] for i in range(20,26)) / (fight_time * 1000)
		might_25_uptime = might_stacks[25] / (fight_time * 1000)

		output_string = '|'+name+' |'+' {{'+prof+'}} | '+my_value(round(fight_time))
		output_string += '|'+"{:.2f}".format(avg_might)
		output_string += "| "+"{:.2f}".format(round((might_uptime * 100), 4))+"%"
		output_string += "| "+"{:.2f}".format(round((might_5_uptime * 100), 4))+"%"
		output_string += "| "+"{:.2f}".format(round((might_10_uptime * 100), 4))+"%"
		output_string += "| "+"{:.2f}".format(round((might_15_uptime * 100), 4))+"%"
		output_string += "| "+"{:.2f}".format(round((might_20_uptime * 100), 4))+"%"
		output_string += "| "+"{:.2f}".format(round((might_25_uptime * 100), 4))+"%"
		output_string += '|'

		print_to_file(output, output_string)

	print_to_file(output, "</$reveal>\n")
	
	# Stability stack table
	print_to_file(output, '<$reveal type="match" state="$:/state/curStackingBuffs" text="stability">\n')
	print_to_file(output, '\n---\n')
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|{{Stability}} uptime by stack|c")
	print_to_file(output, '|thead-dark table-hover sortable|k')
	output_header =  '|!Name | !Class'
	output_header += ' | ! <span data-tooltip="Number of seconds player was in squad logs">Seconds</span>'
	output_header += '| !Avg| !1+ %| !2+ %| !5+ %'
	output_header += '|h'
	print_to_file(output, output_header)
	
	stability_sorted_stacking_uptime_Table = []
	for uptime_prof_name in stacking_uptime_Table:
		fight_time = (stacking_uptime_Table[uptime_prof_name]['duration_stability'] / 1000) or 1
		stability_stacks = stacking_uptime_Table[uptime_prof_name]['stability']

		if (DPSStats[uptime_prof_name]['duration'] * 100) / max_fightTime < config.min_attendance_percentage_for_top:
			continue

		avg_stab = sum(stack_num * stability_stacks[stack_num] for stack_num in range(1, 26)) / (fight_time * 1000)
		stability_sorted_stacking_uptime_Table.append([uptime_prof_name, avg_stab])
	stability_sorted_stacking_uptime_Table = sorted(stability_sorted_stacking_uptime_Table, key=lambda x: x[1], reverse=True)
	stability_sorted_stacking_uptime_Table = list(map(lambda x: x[0], stability_sorted_stacking_uptime_Table))
	
	for uptime_prof_name in stability_sorted_stacking_uptime_Table:
		name = stacking_uptime_Table[uptime_prof_name]['name']
		prof = stacking_uptime_Table[uptime_prof_name]['profession']
		fight_time = (stacking_uptime_Table[uptime_prof_name]['duration_stability'] / 1000) or 1
		stability_stacks = stacking_uptime_Table[uptime_prof_name]['stability']

		avg_stab = sum(stack_num * stability_stacks[stack_num] for stack_num in range(1, 26)) / (fight_time * 1000)
		stab_uptime = 1.0 - (stability_stacks[0] / (fight_time * 1000))
		stab_2_uptime = sum(stability_stacks[i] for i in range(2,26)) / (fight_time * 1000)
		stab_5_uptime = sum(stability_stacks[i] for i in range(5,26)) / (fight_time * 1000)

		output_string = '|'+name+' |'+' {{'+prof+'}} | '+my_value(round(fight_time))
		output_string += '|'+"{:.2f}".format(avg_stab)
		output_string += "| "+"{:.2f}".format(round((stab_uptime * 100), 4))+"%"
		output_string += "| "+"{:.2f}".format(round((stab_2_uptime * 100), 4))+"%"
		output_string += "| "+"{:.2f}".format(round((stab_5_uptime * 100), 4))+"%"
		output_string += '|'

		print_to_file(output, output_string)

	print_to_file(output, "</$reveal>\n")
	print_to_file(output, "</$reveal>\n")
	
	write_stacking_buff_uptimes_in_xls(stacking_uptime_Table, args.xls_output_filename)
	#end Stacking Buff Uptime Table insert


	#start Stacking Buff Uptime Table insert
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Damage with Buffs">')    
	print_to_file(output, '\n<<alert dark "Damage with Buffs" width:60%>>\n\n')
	print_to_file(output, '\n---\n')
	print_to_file(output, '!!! `Damage with buff %` \n')
	print_to_file(output, '!!! Percentage of damage done with a buff, similar to uptime %, but based on damage dealt \n')
	print_to_file(output, '!!! `Damage % - Uptime %` \n')
	print_to_file(output, '!!! The difference in `damage with buff %` and `uptime %` \n')
	print_to_file(output, '\n---\n')
	print_to_file(output, '<$button setTitle="$:/state/curDamageWithBuffs" setTo="might" selectedClass="" class="btn btn-sm btn-dark" style="">might</$button>')
	print_to_file(output, '<$button setTitle="$:/state/curDamageWithBuffs" setTo="other" selectedClass="" class="btn btn-sm btn-dark" style="">other buffs</$button>')
	
	print_to_file(output, '\n---\n')

	dps_sorted_stacking_uptime_Table = []
	for uptime_prof_name in stacking_uptime_Table:
		dps_sorted_stacking_uptime_Table.append([uptime_prof_name, DPSStats[uptime_prof_name]['Damage_Total'] / DPSStats[uptime_prof_name]['duration']])
	dps_sorted_stacking_uptime_Table = sorted(dps_sorted_stacking_uptime_Table, key=lambda x: x[1], reverse=True)
	dps_sorted_stacking_uptime_Table = list(map(lambda x: x[0], dps_sorted_stacking_uptime_Table))

	# Might
	print_to_file(output, '<$reveal type="match" state="$:/state/curDamageWithBuffs" text="might">\n')
	print_to_file(output, '\n---\n')

	# Might with damage table
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|{{Might}} Sortable table - Click header item to sort table {{Might}}|c")
	print_to_file(output, '|thead-dark table-hover sortable|k')
	output_header =  '|!Name | !Class | !DPS' 
	output_header += ' | ! <span data-tooltip="Number of seconds player was in squad logs">Seconds</span>'
	output_header += '| !Avg| !1+ %| !5+ %| !10+ %| !15+ %| !20+ %| !25 %'
	output_header += '|h'
	print_to_file(output, output_header)
	
	for uptime_prof_name in dps_sorted_stacking_uptime_Table:
		name = stacking_uptime_Table[uptime_prof_name]['name']
		prof = stacking_uptime_Table[uptime_prof_name]['profession']
		fight_time = (stacking_uptime_Table[uptime_prof_name]['duration_might'] / 1000) or 1
		damage_with_might = stacking_uptime_Table[uptime_prof_name]['damage_with_might']
		might_stacks = stacking_uptime_Table[uptime_prof_name]['might']

		if stacking_uptime_Table[uptime_prof_name]['duration_might'] * 10 < max_stacking_buff_fight_time:
			continue

		total_damage = DPSStats[uptime_prof_name]["Damage_Total"] or 1
		playerDPS = total_damage/DPSStats[uptime_prof_name]['duration']

		damage_with_avg_might = sum(stack_num * damage_with_might[stack_num] for stack_num in range(1, 26)) / total_damage
		damage_with_might_uptime = 1.0 - (damage_with_might[0] / total_damage)
		damage_with_might_5_uptime = sum(damage_with_might[i] for i in range(5,26)) / total_damage
		damage_with_might_10_uptime = sum(damage_with_might[i] for i in range(10,26)) / total_damage
		damage_with_might_15_uptime = sum(damage_with_might[i] for i in range(15,26)) / total_damage
		damage_with_might_20_uptime = sum(damage_with_might[i] for i in range(20,26)) / total_damage
		damage_with_might_25_uptime = damage_with_might[25] / total_damage
		
		avg_might = sum(stack_num * might_stacks[stack_num] for stack_num in range(1, 26)) / (fight_time * 1000)
		might_uptime = 1.0 - (might_stacks[0] / (fight_time * 1000))
		might_5_uptime = sum(might_stacks[i] for i in range(5,26)) / (fight_time * 1000)
		might_10_uptime = sum(might_stacks[i] for i in range(10,26)) / (fight_time * 1000)
		might_15_uptime = sum(might_stacks[i] for i in range(15,26)) / (fight_time * 1000)
		might_20_uptime = sum(might_stacks[i] for i in range(20,26)) / (fight_time * 1000)
		might_25_uptime = might_stacks[25] / (fight_time * 1000)


		output_string = '|'+name+' |'+' {{'+prof+'}} | '+my_value(round(playerDPS))+'| '+my_value(round(fight_time))

		output_string += '| <span data-tooltip="'+"{:.2f}".format(round(damage_with_avg_might, 4))+'% dmg - '+"{:.2f}".format(round(avg_might, 4))+'% uptime">'
		output_string += "{:.2f}".format(round((damage_with_avg_might), 4))+'</span>'

		output_string += '| <span data-tooltip="'+"{:.2f}".format(round(damage_with_might_uptime * 100, 4))+'% dmg - '+"{:.2f}".format(round(might_uptime * 100, 4))+'% uptime">'
		output_string += "{:.2f}".format(round((damage_with_might_uptime * 100), 4))+'</span>'

		output_string += '| <span data-tooltip="'+"{:.2f}".format(round(damage_with_might_5_uptime * 100, 4))+'% dmg - '+"{:.2f}".format(round(might_5_uptime * 100, 4))+'% uptime">'
		output_string += "{:.2f}".format(round((damage_with_might_5_uptime * 100), 4))+'</span>'

		output_string += '| <span data-tooltip="'+"{:.2f}".format(round(damage_with_might_10_uptime * 100, 4))+'% dmg - '+"{:.2f}".format(round(might_10_uptime * 100, 4))+'% uptime">'
		output_string += "{:.2f}".format(round((damage_with_might_10_uptime * 100), 4))+'</span>'

		output_string += '| <span data-tooltip="'+"{:.2f}".format(round(damage_with_might_15_uptime * 100, 4))+'% dmg - '+"{:.2f}".format(round(might_15_uptime * 100, 4))+'% uptime">'
		output_string += "{:.2f}".format(round((damage_with_might_15_uptime * 100), 4))+'</span>'

		output_string += '| <span data-tooltip="'+"{:.2f}".format(round(damage_with_might_20_uptime * 100, 4))+'% dmg - '+"{:.2f}".format(round(might_20_uptime * 100, 4))+'% uptime">'
		output_string += "{:.2f}".format(round((damage_with_might_20_uptime * 100), 4))+'</span>'

		output_string += '| <span data-tooltip="'+"{:.2f}".format(round(damage_with_might_25_uptime * 100, 4))+'% dmg - '+"{:.2f}".format(round(might_25_uptime * 100, 4))+'% uptime">'
		output_string += "{:.2f}".format(round((damage_with_might_25_uptime * 100), 4))+'</span>'
		
		output_string += '|'

		print_to_file(output, output_string)

	print_to_file(output, "</$reveal>\n")

	# Other buffs with damage
	print_to_file(output, '<$reveal type="match" state="$:/state/curDamageWithBuffs" text="other">\n')
	print_to_file(output, '\n---\n')
		
	# Other buffs with damage table
	other_buffs_with_damage = ['stability', 'protection', 'aegis', 'fury', 'resistance', 'resolution', 'quickness', 'swiftness', 'alacrity', 'vigor', 'regeneration']
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|Sortable table - Click header item to sort table |c")
	print_to_file(output, '|thead-dark table-hover sortable|k')
	output_header =  '|!Name | !Class | !DPS '
	output_header += ' | ! <span data-tooltip="Number of seconds player was in squad logs">Seconds</span>'
	for damage_buff in other_buffs_with_damage:
		output_header += '| !{{'+damage_buff.capitalize()+'}}'
	output_header += '|h'
	print_to_file(output, output_header)
	
	for uptime_prof_name in dps_sorted_stacking_uptime_Table:
		name = stacking_uptime_Table[uptime_prof_name]['name']
		prof = stacking_uptime_Table[uptime_prof_name]['profession']
		uptime_table_prof_name = "{{"+prof+"}} "+name
		if uptime_table_prof_name in uptime_Table:
			uptime_fight_time = uptime_Table[uptime_table_prof_name]['duration'] or 1
		else:
			uptime_fight_time = 1
		dps_fight_time = DPSStats[uptime_prof_name]['duration'] or 1
		fight_time = (stacking_uptime_Table[uptime_prof_name]['duration_might'] / 1000) or 1

		if stacking_uptime_Table[uptime_prof_name]['duration_might'] * 10 < max_stacking_buff_fight_time:
			continue

		total_damage = DPSStats[uptime_prof_name]["Damage_Total"] or 1
		playerDPS = total_damage/dps_fight_time
		output_string = '|'+name+' |'+' {{'+prof+'}} | '+my_value(round(playerDPS))+'| '+my_value(round(fight_time))+'|'

		for damage_buff in other_buffs_with_damage:
			damage_with_buff = stacking_uptime_Table[uptime_prof_name]['damage_with_'+damage_buff]
			damage_with_buff_uptime = damage_with_buff[1] / total_damage			

			if damage_buff in uptime_Table[uptime_table_prof_name]:
				buff_uptime = uptime_Table[uptime_table_prof_name][damage_buff] / uptime_fight_time
			else:
				buff_uptime = 0

			output_string += ' <span data-tooltip="'+"{:.2f}".format(round(damage_with_buff_uptime * 100, 4))+'% dmg - '+"{:.2f}".format(round(buff_uptime * 100, 4))+'% uptime">'
			output_string += "{:.2f}".format(round((damage_with_buff_uptime * 100), 4))+'</span>|'

		print_to_file(output, output_string)

	print_to_file(output, "</$reveal>\n")

	print_to_file(output, "</$reveal>\n")
	#end Stacking Buff Uptime Table insert


	#start On Tag Death insert
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="On Tag Review">')    
	print_to_file(output, '\n<<alert dark "On Tag Review" width:60%>>\n\n')
	print_to_file(output, '\nAvg Dist calculation stops on initial player death or Tag Death to avoiding respawn range')
	print_to_file(output, '\nOn Tag Death Review Current Formula: (On Tag <= 600 Range, Off Tag >600 and <=5000, Run Back Death > 5000)\n')
	print_to_file(output, '\n---\n')
	print_to_file(output, '\n---\n')

	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|Sortable table - Click header item to sort table |c")
	print_to_file(output, "|thead-dark table-hover sortable|k")
	print_to_file(output, "|!Name | !Profession | !Attendance | !Avg Dist| !On-Tag<br>{{Deaths}} |  !Off-Tag<br>{{Deaths}} | !After-Tag<br>{{Deaths}} |  !Run-Back<br>{{Deaths}} |  !Total<br>{{Deaths}} | Off-Tag Deaths Ranges |h")
	for deathOnTag_prof_name in Death_OnTag:
		name = Death_OnTag[deathOnTag_prof_name]['name']
		prof = Death_OnTag[deathOnTag_prof_name]['profession']
		fightTime = uptime_Table.get(deathOnTag_prof_name, {}).get('duration', 1)
		if len(Death_OnTag[deathOnTag_prof_name]["distToTag"]):
			Avg_Dist = round(sum(Death_OnTag[deathOnTag_prof_name]["distToTag"])/len(Death_OnTag[deathOnTag_prof_name]["distToTag"]))
		else:
			Avg_Dist = "N/A"

		if Death_OnTag[deathOnTag_prof_name]['Off_Tag']:
			converted_list = [str(round(element)) for element in Death_OnTag[deathOnTag_prof_name]['Ranges']]
			Ranges_string = ",".join(converted_list)
		else:
			Ranges_string = " "

		output_string = "|"+name+" |"
		output_string += " {{"+prof+"}} | "+my_value(round(fightTime))+" | "+str(Avg_Dist)+"| "+str(Death_OnTag[deathOnTag_prof_name]['On_Tag'])+" | "+str(Death_OnTag[deathOnTag_prof_name]['Off_Tag'])+" | "+str(Death_OnTag[deathOnTag_prof_name]['After_Tag_Death'])+" | "+str(Death_OnTag[deathOnTag_prof_name]['Run_Back'])+" | "+str(Death_OnTag[deathOnTag_prof_name]['Total'])+" |"+Ranges_string+" |"
	


		print_to_file(output, output_string)

	write_Death_OnTag_xls(Death_OnTag, uptime_Table, players, args.xls_output_filename)
	print_to_file(output, "</$reveal>\n")
	#end On Tag Death insert

	#Downed Healing
	down_Heal_Order = {
		55026: 'Glyph of Stars - CA', 69336:"Nature's Renewal", 99999: "Search And Rescue", 14419: 'Battle Standard',
		9163: 'Signet of Mercy', 5763: 'Glyph of Renewal', 10611: 'Signet of Undeath', 59510: "Life Transfer", 
		10527: "Well of Blood", 30504: "Soul Spiral", 6091: "Toss Elixer R", 9999: "Function Gyro", 13117: "Shadow Refuge",
		}
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Downed_Healing">')    
	print_to_file(output, '\n<<alert dark "Healing to downed players" width:60%>>\n\n')
	print_to_file(output, '\nRequires Heal Stat addon for ARCDPS to track\n')
	print_to_file(output, '\n---\n')
	print_to_file(output, '\n---\n')

	print_to_file(output, '\n<div class="flex-row">\n<div class="flex-col border">\n')
	print_to_file(output, "\n!!Healing done\nWork in Progress more skills to be added when logs available\n")
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|Sortable table - Click header item to sort table |c")
	print_to_file(output, "|thead-dark table-hover sortable|k")
	output_string = "|!Name | !Prof | !Combat Time |"
	for key, value in down_Heal_Order.items():
		output_string += "!{{"+value+"}}|"
	output_string += "h"
	print_to_file(output, output_string)
	for squadDps_prof_name in downed_Healing:
		name = downed_Healing[squadDps_prof_name]['name']
		prof = downed_Healing[squadDps_prof_name]['prof']
		fightTime = uptime_Table[squadDps_prof_name]['duration']

		output_string = "|"+name+" | {{"+prof+"}} | "+my_value(round(fightTime))
		for skill in down_Heal_Order:
			if down_Heal_Order[skill] in downed_Healing[squadDps_prof_name]:
				output_string += " | "+my_value(downed_Healing[squadDps_prof_name][down_Heal_Order[skill]]['Heals'])
			else:
				output_string += " | "
		output_string +="|"
		print_to_file(output, output_string)
	
	print_to_file(output, '\n</div>\n<div class="flex-col border">\n')
	print_to_file(output, "\n!!Number of Skill Hits\nWork in Progress more skills to be added when logs available\n")
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|Sortable table - Click header item to sort table |c")
	print_to_file(output, "|thead-dark table-hover sortable|k")
	output_string = "|!Name | !Prof | !Combat Time |"
	for item in down_Heal_Order:
		output_string += "!{{"+down_Heal_Order[item]+"}}|"
	output_string += "h"
	print_to_file(output, output_string)
	for squadDps_prof_name in downed_Healing:
		name = downed_Healing[squadDps_prof_name]['name']
		prof = downed_Healing[squadDps_prof_name]['prof']
		fightTime = uptime_Table[squadDps_prof_name]['duration']

		output_string = "|"+name+" | {{"+prof+"}} | "+my_value(round(fightTime))
		for skill in down_Heal_Order:
			if down_Heal_Order[skill] in downed_Healing[squadDps_prof_name]:
				output_string += " |"+my_value(downed_Healing[squadDps_prof_name][down_Heal_Order[skill]]['Hits'])
			else:
				output_string += " | "
		output_string += "|"
		print_to_file(output, output_string)



	print_to_file(output, '\n</div>\n</div>\n</$reveal>\n')
	#End Downed Healing

	#start Offensive Stat Table insert
	offensive_Order = ['Critical',  'Flanking',  'Glancing',  'Moving',  'Blinded',  'Interupt',  'Invulnerable',  'Evaded',  'Blocked', 'Hard CC', 'CC Duration']
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Offensive Stats">')    
	print_to_file(output, '\n<<alert dark "Offensive Stats across all fights attended." width:60%>>\n\n')
	
	print_to_file(output, '\n---\n')
	print_to_file(output, '\n---\n')

	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|Sortable table - Click header item to sort table |c")
	print_to_file(output, "|thead-dark table-hover sortable|k")
	print_to_file(output, "|!Name | !Profession | !{{Critical}}% |  !{{Flanking}}% |  !{{Glancing}}% |  !{{Moving}}% |  !{{Blind}} |  !{{Interupt}} |  !{{Invulnerable}} |  !{{Evaded}} |  !{{Blocked}} |h")
	for squadDps_prof_name in squad_offensive:
		name = squad_offensive[squadDps_prof_name]['name']
		prof = squad_offensive[squadDps_prof_name]['prof']

		output_string = "|"+name+" | {{"+prof+"}} | "

		#Calculate Critical_Hits_Rate
		if squad_offensive[squadDps_prof_name]['stats']['criticalRate'] and squad_offensive[squadDps_prof_name]['stats']['critableDirectDamageCount']:
			Critical_Rate = round((squad_offensive[squadDps_prof_name]['stats']['criticalRate']/squad_offensive[squadDps_prof_name]['stats']['critableDirectDamageCount'])*100, 4)
		else:
			Critical_Rate = 0.0000
		Critical_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['criticalRate'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['critableDirectDamageCount'])+' critable hits">'+str(Critical_Rate)+'</span>'
		
		output_string += str(Critical_Rate_TT)+" | "
		
		#Calculate Flanking_Rate
		if squad_offensive[squadDps_prof_name]['stats']['flankingRate'] and squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount']:
			Flanking_Rate = round((squad_offensive[squadDps_prof_name]['stats']['flankingRate']/squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])*100, 4)
		else:
			Flanking_Rate = 0.0000
		Flanking_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['flankingRate'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])+' connected direct hit(s)">'+str(Flanking_Rate)+'</span>'
		
		output_string += str(Flanking_Rate_TT)+" | "
		
		#Calculate Glancing Rate
		if squad_offensive[squadDps_prof_name]['stats']['glanceRate'] and squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount']:
			Glancing_Rate = round((squad_offensive[squadDps_prof_name]['stats']['glanceRate']/squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])*100, 4)
		else:
			Glancing_Rate = 0.0000
		Glancing_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['glanceRate'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])+' connected direct hit(s)">'+str(Glancing_Rate)+'</span>'
		
		output_string += str(Glancing_Rate_TT)+" | "
		
		#Calculate Moving_Rate
		if squad_offensive[squadDps_prof_name]['stats']['againstMovingRate'] and squad_offensive[squadDps_prof_name]['stats']['totalDamageCount']:
			Moving_Rate = round((squad_offensive[squadDps_prof_name]['stats']['againstMovingRate']/squad_offensive[squadDps_prof_name]['stats']['totalDamageCount'])*100, 4)
		else:
			Moving_Rate = 0.0000
		Moving_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['againstMovingRate'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['totalDamageCount'])+' direct hit(s)">'+str(Moving_Rate)+'</span>'
		
		output_string += str(Moving_Rate_TT)+" | "
		
		#Calculate Blinded_Rate
		if squad_offensive[squadDps_prof_name]['stats']['missed']:
			Blinded_Rate = squad_offensive[squadDps_prof_name]['stats']['missed']
		else:
			Blinded_Rate = 0
		Blinded_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['missed'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['totalDamageCount'])+' direct hit(s)">'+str(Blinded_Rate)+'</span>'
		
		output_string += str(Blinded_Rate_TT)+" | "
		
		#Calculate Interupt_Rate
		if squad_offensive[squadDps_prof_name]['stats']['interrupts']:
			Interupt_Rate = squad_offensive[squadDps_prof_name]['stats']['interrupts']
		else:
			Interupt_Rate = 0		
		Interupt_Rate_TT = '<span data-tooltip="Interupted enemy players '+str(Interupt_Rate)+' time(s)">'+str(Interupt_Rate)+'</span>'
		
		output_string += str(Interupt_Rate_TT)+" | "
		
		#Calculate Invulnerable_Rate
		if squad_offensive[squadDps_prof_name]['stats']['invulned']:
			Invulnerable_Rate = squad_offensive[squadDps_prof_name]['stats']['invulned']
		else:
			Invulnerable_Rate = 0
		Invulnerable_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['invulned'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['totalDamageCount'])+' hit(s)">'+str(Invulnerable_Rate)+'</span>'
		
		output_string += str(Invulnerable_Rate_TT)+" | "
		
		#Calculate Evaded_Rate
		if squad_offensive[squadDps_prof_name]['stats']['evaded']:
			Evaded_Rate = squad_offensive[squadDps_prof_name]['stats']['evaded']
		else:
			Evaded_Rate = 0
		Evaded_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['evaded'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])+' direct hit(s)">'+str(Evaded_Rate)+'</span>'
		
		output_string += str(Evaded_Rate_TT)+" | "
		
		#Calculate Blocked_Rate
		if squad_offensive[squadDps_prof_name]['stats']['blocked']:
			Blocked_Rate = squad_offensive[squadDps_prof_name]['stats']['blocked']
		else:
			Blocked_Rate = 0		
		Blocked_Rate_TT = '<span data-tooltip="'+str(squad_offensive[squadDps_prof_name]['stats']['blocked'])+' out of '+str(squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])+' direct hit(s)">'+str(Blocked_Rate)+'</span>'
		
		output_string += str(Blocked_Rate_TT)+" |"
		
		print_to_file(output, output_string)

	write_squad_offensive_xls(squad_offensive, args.xls_output_filename)
	print_to_file(output, "</$reveal>\n")
	#end Offensive Stat Table insert

	#start Damage Overview Table insert
	top_players_Totals = get_top_players(players, config, 'dmg', StatType.TOTAL)
	top_players_Averages = get_top_players(players, config, 'dmg', StatType.AVERAGE)
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Damage Overview">')    
	print_to_file(output, '\n<<alert dark "Damage Stats across all fights attended." width:60%>>\n\n')
	
	print_to_file(output, '\n---\n')
	

	print_to_file(output, '<$button setTitle="$:/state/DmgTable" setTo="Total" class="btn btn-sm btn-dark"> Total Damage Stats </$button>')
	print_to_file(output, '<$button setTitle="$:/state/DmgTable" setTo="Average" class="btn btn-sm btn-dark"> Average Damage Stats </$button>\n')
	
	print_to_file(output, '<$reveal type="match" state="$:/state/DmgTable" text="Total">')
	print_to_file(output, '\n---\n')
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "| Sortable Total Stats |c")
	print_to_file(output, "|thead-dark table-hover sortable|k")
	header = "|Name | Profession | Duration"
	for stat in DmgOverviewTable:
		header +="| !"+DmgOverviewTable[stat]
	header+="|h"
	print_to_file(output, header)
	for i in range(len(top_players_Totals)):
		player = players[top_players_Totals[i]]
	#for player in players:
		name = player.name
		prof = player.profession
		durationActive = player.duration_fights_present
		details = "|"+name+" | {{"+prof+"}} | "+my_value(durationActive)
		for stat in DmgOverviewTable:
			if stat == 'dcPct':
				if player.total_stats['dmg'] == 0:
					curStat = 0.0
					details += "| "+my_value(curStat)+"%"
				else:
					curStat = round((player.total_stats['downContribution']/player.total_stats['dmg'])*100, 1)
					details += "| "+my_value(curStat)+"%"
			else:
				curStat = round(player.total_stats[stat], 1)
				details += "| "+my_value(curStat)
		details += "|"
		print_to_file(output, details)
	print_to_file(output, "</$reveal>\n")
	
	print_to_file(output, '<$reveal type="match" state="$:/state/DmgTable" text="Average">')
	print_to_file(output,"\n---\n")
	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "| Sortable Average Stats |c")
	print_to_file(output, "|thead-dark table-hover sortable|k")
	header = "|Name | Profession | Duration"
	for stat in DmgOverviewTable:
		header +="| !"+DmgOverviewTable[stat]
	header+="|h"
	print_to_file(output, header)
	for i in range(len(top_players_Averages)):
		player = players[top_players_Averages[i]]
	#for player in players:
		name = player.name
		prof = player.profession
		durationActive = player.duration_fights_present
		details = "|"+name+" | {{"+prof+"}} | "+my_value(durationActive)
		for stat in DmgOverviewTable:
			if stat == 'dcPct':
				if player.average_stats['dmg'] == 0:
					curStat = 0.0
					details += "| "+my_value(curStat)+"%"
				else:
					curStat = round((player.average_stats['downContribution']/player.average_stats['dmg'])*100, 1)
					details += "| "+my_value(curStat)+"%"
			else:
				curStat = round(player.average_stats[stat], 3)
				details += "| "+"{:,.2f}".format(curStat)
		details += "|"
		print_to_file(output, details)
	print_to_file(output, "</$reveal>\n")

	print_to_file(output, "</$reveal>\n")
	#end Damage Overview Table insert

	# Firebrand pages
	tome1_skill_ids = ["41258", "40635", "42449", "40015", "42898"]
	tome2_skill_ids = ["45022", "40679", "45128", "42008", "42925"]
	tome3_skill_ids = ["42986", "41968", "41836", "40988", "44455"]
	tome_skill_ids = [
		*tome1_skill_ids,
		*tome2_skill_ids,
		*tome3_skill_ids,
	]

	tome_skill_page_cost = {
		"41258": 1, "40635": 1, "42449": 1, "40015": 1, "42898": 1,
		"45022": 1, "40679": 1, "45128": 1, "42008": 2, "42925": 2,
		"42986": 1, "41968": 1, "41836": 2, "40988": 2, "44455": 2,
	}
	
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="FBPages">\n')    
	print_to_file(output, '\n<<alert dark "Firebrand Pages" width:60%>>\n\n')

	print_to_file(output, "|table-caption-top|k")
	print_to_file(output, "|Firebrand page utilization, pages/minute|c")
	print_to_file(output, '|thead-dark table-hover sortable|k')
	output_header =  '|!Name '
	output_header += ' | ! <span data-tooltip="Number of seconds player was in squad logs">Seconds</span>'
	output_header += '| !Pages/min| | !T1 {{Tome_of_Justice}}| !C1 {{Chapter_1_Searing_Spell}}| !C2 {{Chapter_2_Igniting_Burst}}| !C3 {{Chapter_3_Heated_Rebuke}}| !C4 {{Chapter_4_Scorched_Aftermath}}| !Epi {{Epilogue_Ashes_of_the_Just}}| | !T2 {{Tome_of_Resolve}} | !C1 {{Chapter_1_Desert_Bloom}}| !C2 {{Chapter_2_Radiant_Recovery}}| !C3 {{Chapter_3_Azure_Sun}}| !C4 {{Chapter_4_Shining_River}}| !Epi {{Epilogue_Eternal_Oasis}}|  | !T3 {{Tome_of_Courage}}| !C1 {{Chapter_1_Unflinching_Charge}}| !C2 {{Chapter_2_Daring_Challenge}}| !C3 {{Chapter_3_Valiant_Bulwark}}| !C4 {{Chapter_4_Stalwart_Stand}}| !Epi {{Epilogue_Unbroken_Lines}}'
	output_header += '|h'
	print_to_file(output, output_header)
	
	pages_sorted_stacking_uptime_Table = []
	#FB_Pages[player_prof_name]["firebrand_pages"]
	for uptime_prof_name in FB_Pages:
		fight_time = FB_Pages[uptime_prof_name]["fightTime"] or 1
		#stability_stacks = stacking_uptime_Table[uptime_prof_name]['stability']

		#if (DPSStats[uptime_prof_name]['duration'] * 100) / max_fightTime < config.min_attendance_percentage_for_top:
		#	continue

		firebrand_pages = FB_Pages[uptime_prof_name]['firebrand_pages']
		
		all_tomes_total = 0
		for skill_id in tome_skill_ids:
			all_tomes_total += firebrand_pages.get(skill_id, 0) * tome_skill_page_cost[skill_id]

		pages_sorted_stacking_uptime_Table.append([uptime_prof_name, all_tomes_total / fight_time])
	pages_sorted_stacking_uptime_Table = sorted(pages_sorted_stacking_uptime_Table, key=lambda x: x[1], reverse=True)
	pages_sorted_stacking_uptime_Table = list(map(lambda x: x[0], pages_sorted_stacking_uptime_Table))

	def fmt_firebrand_page_total(page_casts, page_cost, fight_time, page_total):
		output_string = ' <span data-tooltip="'

		if page_cost:
			output_string += "{:.2f}".format(round(100 * page_casts * page_cost / page_total, 4))
			output_string += '% of total pages '
			output_string += "{:.2f}".format(round(60 * page_casts / fight_time, 4))
			output_string += ' casts / minute">'
		else:
			output_string += "{:.2f}".format(round(100 * page_casts / page_total, 4))
			output_string += '% of total pages">'

		if page_cost:
			output_string += "{:.2f}".format(round(60 * page_casts * page_cost / fight_time, 4))
		else:
			output_string += "{:.2f}".format(round(60 * page_casts / fight_time, 4))

		output_string += '</span>|'

		return output_string

	
	for uptime_prof_name in pages_sorted_stacking_uptime_Table:
		name = FB_Pages[uptime_prof_name]['name']
		#role = FB_Pages[uptime_prof_name]['role']
		fight_time = FB_Pages[uptime_prof_name]["fightTime"] or 1

		firebrand_pages = FB_Pages[uptime_prof_name]['firebrand_pages']
	
		tome1_total = 0
		for skill_id in tome1_skill_ids:
			tome1_total += firebrand_pages.get(skill_id, 0) * tome_skill_page_cost[skill_id]
	
		tome2_total = 0
		for skill_id in tome2_skill_ids:
			tome2_total += firebrand_pages.get(skill_id, 0) * tome_skill_page_cost[skill_id]
	
		tome3_total = 0
		for skill_id in tome3_skill_ids:
			tome3_total += firebrand_pages.get(skill_id, 0) * tome_skill_page_cost[skill_id]
	
		all_tomes_total = tome1_total + tome2_total + tome3_total

		if all_tomes_total == 0:
			continue

		output_string = '|'+name
		#if role != "Support":
		#	output_string += ' (' + role + ')'
		output_string += ' | ' + my_value(round(fight_time))+' | '
		output_string += "{:.2f}".format(round(60 * all_tomes_total / fight_time, 4)) + '|'
		output_string += ' |'

		output_string += fmt_firebrand_page_total(tome1_total, 0, fight_time, all_tomes_total)
		for skill_id in tome1_skill_ids:
			page_total = firebrand_pages.get(skill_id, 0)
			page_cost = tome_skill_page_cost[skill_id]
			output_string += fmt_firebrand_page_total(page_total, page_cost, fight_time, all_tomes_total)
		output_string += " |"

		output_string += fmt_firebrand_page_total(tome2_total, 0, fight_time, all_tomes_total)
		for skill_id in tome2_skill_ids:
			page_total = firebrand_pages.get(skill_id, 0)
			page_cost = tome_skill_page_cost[skill_id]
			output_string += fmt_firebrand_page_total(page_total, page_cost, fight_time, all_tomes_total)
		output_string += " |"

		output_string += fmt_firebrand_page_total(tome3_total, 0, fight_time, all_tomes_total)
		for skill_id in tome3_skill_ids:
			page_total = firebrand_pages.get(skill_id, 0)
			page_cost = tome_skill_page_cost[skill_id]
			output_string += fmt_firebrand_page_total(page_total, page_cost, fight_time, all_tomes_total)

		print_to_file(output, output_string)

	print_to_file(output, "</$reveal>\n")
	#End Firebrand pages

	#start Dashboard insert
	if config.charts:
		print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Dashboard">')    
		print_to_file(output, '\n<<alert dark "Dashboard for various charts" width:60%>>\n\n')
		Dashboard_Charts = ["Stab vs. HardCC","Kills/Downs/DPS", "Fury/Might/DPS", "Deaths/DamageTaken/DistanceFromTag", "Total Boon Generation", "Cleanses/Heals/BoonScore", "BoonStrips/OutgoingControlScore/DPS", "Profession_DPS_BoxPlot", "Player_DPS_BoxPlot", "Profession_SPS_BoxPlot", "Player_SPS_BoxPlot", "Profession_CPS_BoxPlot", "Player_CPS_BoxPlot", "Profession_HPS_BoxPlot", "Player_HPS_BoxPlot"]
		
		for chart in Dashboard_Charts:
			print_to_file(output, '<$button setTitle="$:/state/curChart" setTo="'+chart+'" selectedClass="" class="btn btn-sm btn-dark" style="">'+chart+' </$button>')
		
		print_to_file(output, '\n---\n')
		

		for chart in Dashboard_Charts:
				print_to_file(output, '<$reveal type="match" state="$:/state/curChart" text="'+chart+'">\n')
				print_to_file(output, '\n---\n')
				print_to_file(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')

				if chart == "Stab vs. HardCC":
					print_to_file(output, "\n!!Stability Uptime versus Hard CC\n")
					print_to_file(output, ",,Bubble Size based on Cc Duration output,,\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_stab_CC_BubbleChartData}} $height="500px" $theme="dark"/>')

				if chart == "Kills/Downs/DPS":
					print_to_file(output, "\n!!Kills / Downs / DPS\n")
					print_to_file(output, ",,Bubble Size based on DPS output,,\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_kills_BubbleChartData}} $height="500px" $theme="dark"/>')

				if chart == "Fury/Might/DPS":
					print_to_file(output, "\n!!Fury / Might / DPS\n")
					print_to_file(output, ",,Bubble Size based on DPS output,,\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_fury_might_BubbleChartData}} $height="500px" $theme="dark"/>')

				if chart == "Deaths/DamageTaken/DistanceFromTag":
					print_to_file(output, "\n!!Deaths / Damage Taken / Distance from Tag\n")
					print_to_file(output, ",,Bubble Size based on Average Distance to Tag,,\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_deaths_BubbleChartData}} $height="500px" $theme="dark"/>')

				if chart == "Total Boon Generation":
					playerCount = len(players)
					calcHeight = str(playerCount*25)
					print_to_file(output, "\n!!Total Boon Generation\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_Total_Boon_Generation_BarChartData}} $height="'+calcHeight+'px" $theme="dark"/>')

				if chart == "Cleanses/Heals/BoonScore":
					print_to_file(output, "\n!!Cleanses / Heals / Boon Score\n")
					print_to_file(output, ",,Bubble Size based on Boon Score = Sum of all average Boon and Aura output,,\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_cleanse_BubbleChartData}} $height="500px" $theme="dark"/>')

				if chart == "BoonStrips/OutgoingControlScore/DPS":
					print_to_file(output, "\n!!Boon Strips / Outgoing Control Score / DPS\n")
					print_to_file(output, ",,Bubble Size based on Control Score = Sum of all outgoing control effects,,\n")
					print_to_file(output, ",,Bubble Size based on DPS output,,\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_rips_BubbleChartData}} $height="500px" $theme="dark"/>')

				#Profession_DPS_BoxPlot
				if chart == "Profession_DPS_BoxPlot":
					print_to_file(output, "\n!!Damage per Second Box Plot by Profession\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_DPS_Profession_Box_PlotChartData}} $height="800px" $theme="dark"/>')

				#Player_DPS_BoxPlot
				if chart == "Player_DPS_BoxPlot":
					print_to_file(output, "\n!!Damage per Second Box Plot by Player\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_DPS_Profession_and_Name_Box_PlotChartData}} $height="800px" $theme="dark"/>')

				#Profession_SPS_BoxPlot
				if chart == "Profession_SPS_BoxPlot":
					print_to_file(output, "\n!!Boon Strip per Second Box Plot by Profession\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_SPS_Profession_Box_PlotChartData}} $height="800px" $theme="dark"/>')

				#Player_SPS_BoxPlot
				if chart == "Player_SPS_BoxPlot":
					print_to_file(output, "\n!!Boon Strip per Second Box Plot by Player\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_SPS_Profession_and_Name_Box_PlotChartData}} $height="800px" $theme="dark"/>')

				#Profession_CPS_BoxPlot
				if chart == "Profession_CPS_BoxPlot":
					print_to_file(output, "\n!!Cleanses per Second Box Plot by Profession\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_CPS_Profession_Box_PlotChartData}} $height="800px" $theme="dark"/>')

				#Player_CPS_BoxPlot
				if chart == "Player_CPS_BoxPlot":
					print_to_file(output, "\n!!Cleanses per Second Box Plot by Player\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_CPS_Profession_and_Name_Box_PlotChartData}} $height="800px" $theme="dark"/>')

				#Profession_HPS_BoxPlot
				if chart == "Profession_HPS_BoxPlot":
					print_to_file(output, "\n!!Heals per Second Box Plot by Profession\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_HPS_Profession_Box_PlotChartData}} $height="800px" $theme="dark"/>')

				#Player_HPS_BoxPlot
				if chart == "Player_HPS_BoxPlot":
					print_to_file(output, "\n!!Heals per Second Box Plot by Player\n")
					print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_HPS_Profession_and_Name_Box_PlotChartData}} $height="800px" $theme="dark"/>')

				print_to_file(output, '\n</div>\n</div>\n')
				print_to_file(output, "</$reveal>\n")

		print_to_file(output, "</$reveal>\n")
	#end Dashboard insert

	#start DPS Stats insert		
	sorted_DPSStats = []
	for DPSStats_prof_name in DPSStats:
		name = DPSStats[DPSStats_prof_name]['name']
		prof = DPSStats[DPSStats_prof_name]['profession']
		fightTime = DPSStats[DPSStats_prof_name]['duration'] or 1

		if DPSStats[DPSStats_prof_name]['Damage_Total'] / fightTime < 250 or (fightTime * 100) / max_fightTime < config.min_attendance_percentage_for_top:
			continue

		sorted_DPSStats.append([DPSStats_prof_name, DPSStats[DPSStats_prof_name]['Damage_Total'] / fightTime])
	sorted_DPSStats = sorted(sorted_DPSStats, key=lambda x: x[1], reverse=True)
	sorted_DPSStats = list(map(lambda x: x[0], sorted_DPSStats))

	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="DPSStats">')    
	print_to_file(output, '\n<<alert dark "Experimental DPS stats" width:60%>>\n\n')
	
	print_to_file(output, '\n---\n')
	print_to_file(output, '!!! `Chunk Damage(t)` [`Ch(t)DPS`] \n')
	print_to_file(output, '!!! Damage done `t` seconds before an enemy goes down \n')
	print_to_file(output, '!!! `Carrior Damage` [`CaDPS`] \n')
	print_to_file(output, '!!! Damage done to down enemies that die \n')
	print_to_file(output, '\n---\n')

	print_to_file(output, '|table-caption-top|k')
	print_to_file(output, '|Sortable table - Click header item to sort table |c')
	print_to_file(output, '|thead-dark table-hover sortable|k')
	output_header =  '|!Name | !Class | !Role'
	output_header += ' | ! <span data-tooltip="Number of seconds player was in squad logs">Seconds</span>'
	output_header += '| !DPS| !Ch2DPS| !Ch4DPS| !Ch8DPS| !CaDPS'
	output_header += '|h'
	print_to_file(output, output_header)
	for DPSStats_prof_name in sorted_DPSStats:
		name = DPSStats[DPSStats_prof_name]['name']
		prof = DPSStats[DPSStats_prof_name]['profession']
		role = DPSStats[DPSStats_prof_name]['role']
		fightTime = DPSStats[DPSStats_prof_name]['duration'] or 1

		output_string = '|'+name+' |'+' {{'+prof+'}} | '+role+' | '+my_value(fightTime)
		output_string += '| '+'<span data-tooltip="'+my_value(DPSStats[DPSStats_prof_name]['Damage_Total'])+' total damage">'+my_value(round(DPSStats[DPSStats_prof_name]['Damage_Total'] / fightTime))+'</span>'
		output_string += '| '+'<span data-tooltip="'+my_value(DPSStats[DPSStats_prof_name]['Chunk_Damage'][2])+' chunk(2) damage">'+my_value(round(DPSStats[DPSStats_prof_name]['Chunk_Damage'][2] / fightTime))+'</span>'
		output_string += '| '+'<span data-tooltip="'+my_value(DPSStats[DPSStats_prof_name]['Chunk_Damage'][4])+' chunk (4) damage">'+my_value(round(DPSStats[DPSStats_prof_name]['Chunk_Damage'][4] / fightTime))+'</span>'
		output_string += '| '+'<span data-tooltip="'+my_value(DPSStats[DPSStats_prof_name]['Chunk_Damage'][8])+' chunk (8) damage">'+my_value(round(DPSStats[DPSStats_prof_name]['Chunk_Damage'][8] / fightTime))+'</span>'
		output_string += '| '+'<span data-tooltip="'+my_value(DPSStats[DPSStats_prof_name]['Carrion_Damage'])+' carrion damage">'+my_value(round(DPSStats[DPSStats_prof_name]['Carrion_Damage'] / fightTime))+'</span>'
		output_string += '|'

		print_to_file(output, output_string)

	write_DPSStats_xls(DPSStats, args.xls_output_filename)
	print_to_file(output, '\n---\n')
	print_to_file(output, "\n!!DPS Stats Bubble Chart\n")
	print_to_file(output, "\n,,Bubble size based on CDPS,,\n") 
	if config.charts:
		print_to_file(output, '<$echarts $text={{'+fileDate.strftime("%Y%m%d%H%M")+'_DPSStats_BubbleChartData}} $height="500px" $theme="dark"/>')
	else:
		print_to_file(output, "\n Charts disabled in config\n")	
	print_to_file(output, "</$reveal>\n")
	#end DPS Stats insert

	# Burst Damage
	print_to_file(output, '<$reveal type="match" state="$:/state/curTab" text="Burst Damage">\n')    
	print_to_file(output, '\n<<alert dark "Experimental DPS stats" width:60%>>\n\n')
	
	print_to_file(output, '---\n')
	print_to_file(output, '!!! `Burst Damage(t)` [`Bur(t)`] \n')
	print_to_file(output, '!!! Maximum damage done over any `t` second interval \n')
	print_to_file(output, '---\n')
	print_to_file(output, '!!! `Ch5Ca Burst Damage(t)` [`Ch5CaBur(t)`] \n')
	print_to_file(output, '!!! Maximum Chunk(5) + Carrion damage done over any `t` second interval \n')
	print_to_file(output, '---\n')

	burst_menu_string = '| '
	burst_menu_string += '<$radio tiddler="$:/temp/BurstDamage" field="curBurstTableDamage" value="Ch5Ca">Ch5Ca Damage</$radio>&nbsp; &nbsp;<$radio tiddler="$:/temp/BurstDamage" field="curBurstTableDamage" value="Damage"> Total Damage</$radio>'
	burst_menu_string += '&nbsp;&nbsp;/&nbsp;&nbsp;'
	burst_menu_string += '<$radio tiddler="$:/temp/BurstDamage" field="curBurstTableType" value="Cumulative">&nbsp;Cumulative</$radio>&nbsp; &nbsp;<$radio tiddler="$:/temp/BurstDamage" field="curBurstTableType" value="PS">&nbsp;PS</$radio>'
	burst_menu_string += ' |c'

	# First the per second version of the table
	print_to_file(output, '<$reveal type="match" stateTitle= "$:/temp/BurstDamage" stateField="curBurstTableDamage" text="Damage">\n')
	print_to_file(output, '<$reveal type="match" stateTitle= "$:/temp/BurstDamage" stateField="curBurstTableType" text="PS">\n')

	print_to_file(output, '|table-caption-top|k')
	print_to_file(output, burst_menu_string)
	print_to_file(output, '|thead-dark table-hover sortable|k')
	
	output_string = '|!Name | !Class | !Role |'

	for i in list(range(1, 6)) + list(range(10, 21, 5)):
		output_string += " !"+str(i)+"s |"
		
	output_string += "h"
	print_to_file(output, output_string)

	for DPSStats_prof_name in sorted_DPSStats:
		name = DPSStats[DPSStats_prof_name]['name']
		prof = DPSStats[DPSStats_prof_name]['profession']
		role = DPSStats[DPSStats_prof_name]['role']
		fightTime = DPSStats[DPSStats_prof_name]['duration']

		output_string = '|'+name+' |'+' {{'+prof+'}} | '+role+' |'
		for i in list(range(1, 6)) + list(range(10, 21, 5)):
			output_string += ' '+my_value(round(DPSStats[DPSStats_prof_name]['Burst_Damage'][i] / i))+'|'
				
		print_to_file(output, output_string)

	print_to_file(output, "\n</$reveal>\n")

	# Next the cumulative version of the table
	print_to_file(output, '<$reveal type="match" stateTitle= "$:/temp/BurstDamage" stateField="curBurstTableType" text="Cumulative">\n')

	print_to_file(output, '|table-caption-top|k')
	print_to_file(output, burst_menu_string)
	print_to_file(output, '|thead-dark table-hover sortable|k')
	
	output_string = '|!Name | !Class | !Role |'

	for i in list(range(1, 6)) + list(range(10, 21, 5)):
		output_string += " !"+str(i)+"s |"
		
	output_string += "h"
	print_to_file(output, output_string)

	for DPSStats_prof_name in sorted_DPSStats:
		name = DPSStats[DPSStats_prof_name]['name']
		prof = DPSStats[DPSStats_prof_name]['profession']
		role = DPSStats[DPSStats_prof_name]['role']
		fightTime = DPSStats[DPSStats_prof_name]['duration'] or 1

		output_string = '|'+name+' |'+' {{'+prof+'}} | '+role+' |'
		for i in list(range(1, 6)) + list(range(10, 21, 5)):
			output_string += ' '+my_value(DPSStats[DPSStats_prof_name]['Burst_Damage'][i])+'|'
				
		print_to_file(output, output_string)

	print_to_file(output, "\n</$reveal>\n")
	print_to_file(output, "\n</$reveal>\n")

	# Ch5Ca Burst Damage
	# First the per second version of the table
	print_to_file(output, '<$reveal type="match" stateTitle= "$:/temp/BurstDamage" stateField="curBurstTableDamage" text="Ch5Ca">\n')
	print_to_file(output, '<$reveal type="match" stateTitle= "$:/temp/BurstDamage" stateField="curBurstTableType" text="PS">\n')

	print_to_file(output, '|table-caption-top|k')
	print_to_file(output, burst_menu_string)
	print_to_file(output, '|thead-dark table-hover sortable|k')
	
	output_string = '|!Name | !Class | !Role |'

	for i in list(range(1, 6)) + list(range(10, 21, 5)):
		output_string += " !"+str(i)+"s |"
		
	output_string += "h"
	print_to_file(output, output_string)

	for DPSStats_prof_name in sorted_DPSStats:
		name = DPSStats[DPSStats_prof_name]['name']
		prof = DPSStats[DPSStats_prof_name]['profession']
		role = DPSStats[DPSStats_prof_name]['role']
		fightTime = DPSStats[DPSStats_prof_name]['duration'] or 1

		output_string = '|'+name+' |'+' {{'+prof+'}} | '+role+' |'
		for i in list(range(1, 6)) + list(range(10, 21, 5)):
			output_string += ' '+my_value(round(DPSStats[DPSStats_prof_name]['Ch5Ca_Burst_Damage'][i] / i))+'|'
				
		print_to_file(output, output_string)

	print_to_file(output, "\n</$reveal>\n")

	# Next the cumulative version of the table
	print_to_file(output, '<$reveal type="match" stateTitle= "$:/temp/BurstDamage" stateField="curBurstTableType" text="Cumulative">\n')

	print_to_file(output, '|table-caption-top|k')
	print_to_file(output, burst_menu_string)
	print_to_file(output, '|thead-dark table-hover sortable|k')
	
	output_string = '|!Name | !Class | !Role |'

	for i in list(range(1, 6)) + list(range(10, 21, 5)):
		output_string += " !"+str(i)+"s |"
		
	output_string += "h"
	print_to_file(output, output_string)

	for DPSStats_prof_name in sorted_DPSStats:
		name = DPSStats[DPSStats_prof_name]['name']
		prof = DPSStats[DPSStats_prof_name]['profession']
		role = DPSStats[DPSStats_prof_name]['role']
		fightTime = DPSStats[DPSStats_prof_name]['duration'] or 1

		output_string = '|'+name+' |'+' {{'+prof+'}} | '+role+' |'
		for i in list(range(1, 6)) + list(range(10, 21, 5)):
			output_string += ' '+my_value(DPSStats[DPSStats_prof_name]['Ch5Ca_Burst_Damage'][i])+'|'
				
		print_to_file(output, output_string)

	print_to_file(output, "\n</$reveal>\n")
	print_to_file(output, "\n</$reveal>\n")

	print_to_file(output, "\n</$reveal>\n")     
	# end Ch5Ca Burst Damage

	top_players_by_stat = top_average_stat_players if config.player_sorting_stat_type == 'average' else top_total_stat_players
	for stat in config.stats_to_compute:
		if damage_overview_only and stat in DmgOverviewTable:
			continue
		if defensive_overview_only and tab in excludeForDefOverview:
			continue
		skip_boxplot_charts = ['deaths', 'iol', 'stealth', 'HiS']
		#boxplot_Stats = ['stability',  'protection', 'aegis', 'might', 'fury', 'resistance', 'resolution', 'quickness', 'swiftness', 'alacrity', 'vigor', 'regeneration', 'res', 'kills', 'downs', 'swaps', 'dmg', 'Pdmg', 'Cdmg', 'rips', 'cleanses', 'superspeed', 'barrierDamage']
		if stat == 'dist':
			write_stats_xls(players, top_percentage_stat_players[stat], stat, args.xls_output_filename)
			if config.charts:
				#write_stats_chart(players, top_percentage_stat_players[stat], stat, myDate, args.input_directory, config)
				write_stats_box_plots(players, top_percentage_stat_players[stat], stat, ProfessionColor, myDate, args.input_directory, config)
		#elif stat == 'dmg_taken':
		#	write_stats_xls(players, top_average_stat_players[stat], stat, args.xls_output_filename)
		#	if config.charts:
		#		#write_stats_chart(players, top_average_stat_players[stat], stat, myDate, args.input_directory, config)
		#		write_stats_box_plots(players, top_average_stat_players[stat], stat, ProfessionColor, myDate, args.input_directory, config)
		elif stat == 'heal' and found_healing:
			write_stats_xls(players, top_players_by_stat[stat], stat, args.xls_output_filename)
			if config.charts:
				#write_stats_chart(players, top_players_by_stat[stat], stat, myDate, args.input_directory, config)
				write_stats_box_plots(players, top_players_by_stat[stat], stat, ProfessionColor, myDate, args.input_directory, config)
		elif stat == 'barrier' and found_barrier:
			write_stats_xls(players, top_players_by_stat[stat], stat, args.xls_output_filename)
			if config.charts:
				#write_stats_chart(players, top_players_by_stat[stat], stat, myDate, args.input_directory, config)
				write_stats_box_plots(players, top_players_by_stat[stat], stat, ProfessionColor, myDate, args.input_directory, config)
		#elif stat == 'deaths':
		#	write_stats_xls(players, top_consistent_stat_players[stat], stat, args.xls_output_filename)
		#	if config.charts:
		#		write_stats_chart(players, top_consistent_stat_players[stat], stat, myDate, args.input_directory, config)
		elif stat not in skip_boxplot_charts:
			write_stats_xls(players, top_players_by_stat[stat], stat, args.xls_output_filename)
			if config.charts:
				#write_stats_chart(players, top_players_by_stat[stat], stat, myDate, args.input_directory, config)
				write_stats_box_plots(players, top_players_by_stat[stat], stat, ProfessionColor, myDate, args.input_directory, config)
		else:
			write_stats_xls(players, top_players_by_stat[stat], stat, args.xls_output_filename)
			if config.charts:
				write_stats_chart(players, top_players_by_stat[stat], stat, myDate, args.input_directory, config)
				#write_stats_box_plots(players, top_players_by_stat[stat], stat, ProfessionColor, myDate, args.input_directory, config)
		if stat == 'rips' or stat == 'cleanses' or stat == 'stability' or stat == 'heal':
			supportCount = write_support_xls(players, top_players_by_stat[stat], stat, args.xls_output_filename, supportCount)

	#write out Bubble Charts and Box_Plots
	if config.charts:
		write_bubble_charts(players, top_players_by_stat[stat], squad_Control, myDate, args.input_directory)
		write_TotalBoon_bar_chart(players, myDate, args.input_directory)
		if include_comp_and_review:
			write_spike_damage_heatmap(squad_damage_output, myDate, args.input_directory)
		write_box_plot_charts(DPS_List, myDate, args.input_directory, "DPS")
		write_box_plot_charts(SPS_List, myDate, args.input_directory, "SPS")
		write_box_plot_charts(CPS_List, myDate, args.input_directory, "CPS")
		write_box_plot_charts(HPS_List, myDate, args.input_directory, "HPS")
		write_DPSStats_bubble_charts(uptime_Table, DPSStats, myDate, args.input_directory)
		write_Attendance_xls(Attendance, args.xls_output_filename)

	#check for unknown teamIDs
	if config.check_for_unknown_team_ids:
		print('Checking for unknown teamIDs in all Fights')
		print('FightNum\tTeamID: count')
		teamID_OK = True
		for fight_num, fight in enumerate(fights):
			#check for unknown teamIDs
			if fight.enemies_Unk:
				teamID_OK = False
				print_to_file(log, f'{fight_num+1: >8}\t{str(fight.enemies_Unk)}')
		if teamID_OK:
			print_to_file(log, f'{fight_num+1: >8} fights, No unknown teamIDs')