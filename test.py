# -*- coding: utf-8 -*-
"""
Created on Tue Apr 21 22:41:09 2026

@author: Jeremiah
"""

from datetime import datetime
import xml.etree.ElementTree as ET
from xml.dom import minidom


class TLCParserConfig:
    def __init__(self):
        # Settings defaults
        self.skip_purchased_awards = 1
        self.skip_purchased_ranks = 0
        self.awards_cards_output_file = "awards_detail_cards.html"
        self.awards_program_output_file = "awards_program.html"
        self.awards_shopping_list = ""
        self.full_awards_output_file = "awards_full_listing.html"
        self.merge_with_previous_data = 0
        self.previous_data_file = ""
        self.new_awards_shopping_list = "temp_new_awards_shopping_list.xlsx"

        # DYOBadges: list of dicts
        self.badges = []

    # -------------------------
    # LOAD FROM XML
    # -------------------------
    def load(self, filename):
        tree = ET.parse(filename)
        root = tree.getroot()

        settings = root.find('Settings')
        if settings is not None:
            self.skip_purchased_awards = int(settings.findtext('SkipPurchasedAwards', '1'))
            self.skip_purchased_ranks = int(settings.findtext('SkipPurchasedRanks', '0'))
            self.awards_cards_output_file = settings.findtext('AwardsCardsOutputFile', 'awards_detail_cards.html')
            self.awards_program_output_file = settings.findtext('AwardsProgramOutputFile', 'awards_program.html')
            self.awards_shopping_list = settings.findtext('AwardsShoppingList', '')
            self.full_awards_output_file = settings.findtext('FullAwardsOutputFile', 'awards_full_listing.html')
            self.merge_with_previous_data = int(settings.findtext('MergeWithPreviousData', '0'))
            self.previous_data_file = settings.findtext('PreviousDataFile', '')
            self.new_awards_shopping_list = settings.findtext('NewAwardsShoppingList', 'temp_new_awards_shopping_list.xlsx')
        #if self.awards_shopping_list is None or self.awards_shopping_list=='':
        #    self.awards_shopping_list = 'COH-shopping-list-' + datetime.now().strftime("%Y%m%d%H%M") + '.xlsx'

        # Load badges
        self.badges = []
        dyo = root.find('DYOBadges')
        if dyo is not None:
            for b in dyo.findall('Badge'):
                self.badges.append({
                    'Trailman': b.findtext('Trailman', ''),
                    'CompletedDate': b.findtext('CompletedDate', ''),
                    'BadgeName': b.findtext('BadgeName', '')
                })

    # -------------------------
    # GETTERS
    # -------------------------
    def get_skip_purchased_awards(self):
        return self.skip_purchased_awards

    def get_skip_purchased_ranks(self):
        return self.skip_purchased_ranks

    def get_awards_cards_output_file(self):
        return self.awards_cards_output_file

    def get_awards_program_output_file(self):
        return self.awards_program_output_file

    def get_awards_shopping_list(self):
        return self.awards_shopping_list

    def get_full_awards_output_file(self):
        return self.full_awards_output_file

    def get_merge_with_previous_data(self):
        return self.merge_with_previous_data

    def get_previous_data_file(self):
        return self.previous_data_file

    def get_new_awards_shopping_list(self):
        return self.new_awards_shopping_list

    def get_badges(self):
        return self.badges

    # -------------------------
    # WRITE TO XML (WITH COMMENTS)
    # -------------------------
    def write(self, filename):
        root = ET.Element('TLCParser')

        # ---- Settings ----
        root.append(ET.Comment(
            "Settings group specifies what the parse awards program will output."
        ))
        settings = ET.SubElement(root, 'Settings')

        def add_setting(name, value, comment):
            settings.append(ET.Comment(comment))
            elem = ET.SubElement(settings, name)
            elem.text = str(value)

        add_setting(
            'SkipPurchasedAwards',
            self.skip_purchased_awards,
            "Set to 1 if purchased awards were already announced and packaged or 0 if they should be included in the program."
        )

        add_setting(
            'SkipPurchasedRanks',
            self.skip_purchased_ranks,
            "Set to 1 if purchased ranks were already announced and packaged or 0 if they should be included in the program."
        )

        add_setting(
            'AwardsCardsOutputFile',
            self.awards_cards_output_file,
            "Filename of HTML file containing listing of all awards to be announced and handed out for each Trailman."
        )

        add_setting(
            'AwardsProgramOutputFile',
            self.awards_program_output_file,
            "Filename of HTML file containing the awards program."
        )

        add_setting(
            'AwardsShoppingList',
            self.awards_shopping_list,
            "Filename of XLSX shopping list (blank = default time-stamped name)."
        )

        add_setting(
            'FullAwardsOutputFile',
            self.full_awards_output_file,
            "Optional full awards output file (including purchased awards skipped in the program)."
        )

        add_setting(
            'MergeWithPreviousData',
            self.merge_with_previous_data,
            "Set to 1 to merge with previous data."
        )

        add_setting(
            'PreviousDataFile',
            self.previous_data_file,
            "PKL file to merge with (blank = merge with most recent file)."
        )

        add_setting(
            'NewAwardsShoppingList',
            self.new_awards_shopping_list,
            "Optional shopping list for new awards only (leave blank to skip writing this file)."
        )

        # ---- DYOBadges ----
        if self.badges:
            root.append(ET.Comment(
                """"Optional group listing any custom Design-Your-Own/TEAMS badge names.
                List a separate Badge group for each custom badge. If a badge is found not listed here,
                you will be prompted for a name. The Trailman name (in Lastname, Firstname format) and
                Completion Date (in MM/DD/YYYY format) are used to match the badge name with the correct award."""
            ))
            dyo = ET.SubElement(root, 'DYOBadges')

            for b in self.badges:
                dyo.append(ET.Comment(
                    "Custom badge entry"
                ))
                badge = ET.SubElement(dyo, 'Badge')

                ET.SubElement(badge, 'Trailman').text = b['Trailman']
                ET.SubElement(badge, 'CompletedDate').text = b['CompletedDate']
                ET.SubElement(badge, 'BadgeName').text = b['BadgeName']

        # ---- Pretty print ----
        rough = ET.tostring(root, 'utf-8')
        pretty = minidom.parseString(rough).toprettyxml(indent="  ")

        with open(filename, 'w', encoding='utf-8') as f:
            f.write(pretty)