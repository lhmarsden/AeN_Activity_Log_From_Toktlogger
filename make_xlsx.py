#!/usr/bin/python3
# encoding: utf-8
'''
 -- Creates xlsx files for collecting samples in Nansen Legacy


@author:     Pål Ellingsen, Luke Marsden
@contact:    lukem@unis.no
'''
import xlsxwriter
import pandas as pd
import fields
from argparse import ArgumentParser, RawDescriptionHelpFormatter, Namespace
import os

DEBUG = 1

DEFAULT_FONT = 'Calibri'
DEFAULT_SIZE = 10

class Field(object):
    """
    Object for holding the specification of a cell

    """

    def __init__(self, name, disp_name, validation={},
                 cell_format={}, width=20, long_list=False):
        """
        Initialising the object

        Parameters
        ----------
        name : str
               Name of object

        disp_name : str
               The title of the column

        validation : dict, optional
            A dictionary using the keywords defined in xlsxwriter

        cell_format : dict, optional
            A dictionary using the keywords defined in xlsxwriter

        width : int, optional
            Number of units for width

        long_list : Boolean, optional
            True for enabling the long list.
        """
        self.name = name  # Name of object
        self.disp_name = disp_name  # Title of column
        self.cell_format = cell_format  # For holding the formatting of the cell
        self.validation = validation  # For holding the validation of the cell
        self.long_list = long_list  # For holding the need for an entry in the
        # variables sheet
        self.width = width

    def set_validation(self, validation):
        """
        Set the validation of the cell

        Parameters
        ----------
        validation : dict
            A dictionary using the keywords defined in xlsxwriter

        """
        self.validation = validation

    def set_cell_format(self, cell_format):
        """
        Set the validation of the cell

        Parameters
        ----------
        cell_format : dict
            A dictionary using the keywords defined in xlsxwriter

        """
        self.cell_format = cell_format

    def set_width(self, width):
        """
        Set the cell width

        Parameters
        ----------
        width : int
            Number of units for width


        """
        self.width = width

    def set_long_list(self, long_list):
        """
        Set the need for moving the source in the list to a cell range in the
        variables sheet

        Parameters
        ----------
        long_list : Boolean
            True for enabling the long list.


        """
        self.long_list = long_list


class Variable_sheet(object):
    """
    Class for handling the variable sheet

    """

    def __init__(self, workbook):
        """
        Initialises the sheet

        Parameters
        ----------
        workbook: Xlsxwriter workbook
            The parent workbook where the sheet is added

        """
        self.workbook = workbook
        self.name = 'Variables'  # The name of the worksheet
        self.sheet = workbook.add_worksheet(self.name)
        self.sheet.hide()  # Hide the sheet
        # For holding the current row to add variables on
        self.current_column = 0

    def add_row(self, variable, parameter_list):
        """
        Adds a row of parameters to a variable and returns the ref for the list

        Parameters
        ----------
        variable : str
            The name of the variable

        parameter_list :
            List of parameters to be added

        Returns
        ----------
        ref : str
            The range of the list in Excel format
        """

#         print(parameter_list)
        self.sheet.write(0, self.current_column, variable)
        name = 'Table_' + variable.replace(' ', '_').capitalize()

        tab = self.sheet.add_table(
            1, self.current_column,
            1 + len(parameter_list), self.current_column,
            {'name': name,
                'header_row': 0}
        )
#         print(tab.properties['name'])
        #

        for ii, par in enumerate(sorted(parameter_list, key=str.lower)):
            self.sheet.write(1 + ii, self.current_column, par)
        ref = '=INDIRECT("' + name + '")'

        # Increment row such that the next gets a new row
        self.current_column = self.current_column + 1
        return ref


def make_dict_of_fields():
    """
    Makes a dictionary of the possible fields.
    Does this by reading the fields list from the fields.py library

    Returns
    ---------
    field_dict : dict
        Dictionary of the possible fields
        Contains a Field object under each name


    """

    field_dict = {}
    for field in fields.fields:
        field['name'] = field['name']
        new = Field(field['name'], field['disp_name'])
        if 'valid' in field:
            new.set_validation(field['valid'])
#             print(new.validation)
        if 'cell_format' in field:
            new.set_cell_format(field['cell_format'])
        if 'width' in field:
            new.set_width(field['width'])
        else:
            new.set_width(len(field['disp_name']))
        if 'long_list' in field:
            new.set_long_list(field['long_list'])

        field_dict[field['name']] = new
    return field_dict

def write_conversion(args, workbook):
    """
    Adds a conversion sheet to workbook

    Parameters
    ----------
    args : argparse object
        The input arguments

    workbook : xlsxwriter Workbook
        The workbook for the conversion sheet


    """

    sheet = workbook.add_worksheet('Conversion')

    parameter_format = workbook.add_format({
        'font_name': DEFAULT_FONT,
        'right': True,
        'bottom': True,
        'bold': False,
        'text_wrap': True,
        'valign': 'left',
        'font_size': DEFAULT_SIZE + 2,
        'bg_color': '#B9F6F5',
    })
    center_format = workbook.add_format({
        'font_name': DEFAULT_FONT,
        'right': True,
        'bottom': True,
        'bold': False,
        'text_wrap': True,
        'valign': 'center',
        'font_size': DEFAULT_SIZE + 2,
        'bg_color': '#23EEFF',
    })
    output_format = workbook.add_format({
        'font_name': DEFAULT_FONT,
        'right': True,
        'bottom': True,
        'bold': False,
        'text_wrap': True,
        'valign': 'left',
        'font_size': DEFAULT_SIZE + 2,
        'bg_color': '#FF94E8',
    })

    input_format = workbook.add_format({
        'bold': False,
        'font_name': DEFAULT_FONT,
        'text_wrap': True,
        'valign': 'left',
        'font_size': DEFAULT_SIZE
    })

    height = 15
    sheet.set_column(0, 2, width=30)

    sheet.write(1, 0, "Coordinate conversion ", parameter_format)
    sheet.merge_range(2, 0, 2, 1, "Degree Minutes Seconds ", center_format)
    sheet.write(3, 0, "Degrees ", parameter_format)
    sheet.write(4, 0, "Minutes ", parameter_format)
    sheet.write(5, 0, "Seconds ", parameter_format)
    sheet.write(6, 0, "Decimal degrees ", output_format)
    sheet.write(6, 1, "=B4+B5/60+B6/3600 ", output_format)
    sheet.merge_range(7, 0, 7, 1, "Degree decimal minutes", center_format)
    sheet.write(8, 0, "Degrees ", parameter_format)
    sheet.write(9, 0, "Decimal minutes ", parameter_format)
    sheet.write(10, 0, "Decimal degrees ", output_format)
    sheet.write(10, 1, "=B9+B10/60 ", output_format)


def write_metadata(args, workbook, field_dict, metadata_df):
    """
    Adds a metadata sheet to workbook

    Parameters
    ----------
    args : argparse object
        The input arguments

    workbook : xlsxwriter Workbook
        The workbook for the metadata sheet

    field_dict : dict
        Contains a dictionary of Field objects and their name, made with
        make_dict_of _fields()

    metadata_df: pandas.core.frame.DataFrame
        Optional parameter. Option to add metadata from a dataframe to the 'metadata' sheet.

    """

    sheet = workbook.add_worksheet('Metadata')

    metadata_fields = ['title', 'abstract', 'pi_name', 'pi_email', 'pi_institution',
                       'pi_address', 'recordedBy', 'projectID', 'cruiseNumber', 'vesselName']

    parameter_format = workbook.add_format({
        'font_name': DEFAULT_FONT,
        'right': True,
        'bottom': True,
        'bold': False,
        'text_wrap': True,
        'valign': 'left',
        'font_size': DEFAULT_SIZE + 2,
        'bg_color': '#B9F6F5',
    })
    input_format = workbook.add_format({
        'bold': False,
        'font_name': DEFAULT_FONT,
        'text_wrap': True,
        'valign': 'left',
        'font_size': DEFAULT_SIZE
    })


    sheet.set_column(0, 0, width=30)
    sheet.set_column(2, 2, width=50)
    for ii, mfield in enumerate(metadata_fields):
        field = field_dict[mfield]
        sheet.write(ii, 0, field.disp_name, parameter_format)
        sheet.write(ii, 1, field.name, parameter_format)
        sheet.set_column(1, 1, None, None, {'hidden': True})

        if type(metadata_df) == pd.core.frame.DataFrame:
            try:
                sheet.write(ii,2,metadata_df[mfield][0], input_format)
            except:
                sheet.write(ii, 2, '', input_format)
                continue
        else:
            sheet.write(ii, 2, '', input_format)

        if field.validation:
            if args.verbose > 0:
                print("Writing metadata validation")
            valid_copy = field.validation.copy()
            if len(valid_copy['input_message']) > 255:
                valid_copy['input_message'] = valid_copy[
                    'input_message'][:252] + '...'
            sheet.data_validation(first_row=ii,
                                  first_col=2,
                                  last_row=ii,
                                  last_col=2,
                                  options=valid_copy)
            if field.cell_format:
                cell_format = workbook.add_format(field.cell_format)
                sheet.set_row(
                    ii, ii, cell_format=cell_format)

        if ii == 0:
            height = 30
        elif ii == 1: # Making abstract row height larger to encourage researches to write more.
            height = 150
        else:
            height = 15

        sheet.set_row(ii, height)

def make_xlsx(args, file_def, field_dict, metadata, conversions, data, metadata_df):
    """
    Writes the xlsx file based on the wanted fields

    Parameters
    ----------
    args : argparse object
        The input arguments

    file_def : dict
        The definition of the file wanted, generate this with read_xml

    field_dict : dict
        Contains a dictionary of Field objects and their name, made with
        make_dict_of _fields()

    metadata: Boolean
        Should the metadata sheet be written

    conversions: Boolean
        Should the conversions sheet be written

    data: pandas.core.frame.DataFrame
        Optional parameter. Option to add data from a dataframe to the 'data' sheet.

    metadata_df: pandas.core.frame.DataFrame
        Optional parameter. Option to add metadata from a dataframe to the 'metadata' sheet.

    """

    output = os.path.join(args.dir, file_def['name'] + '.xlsx')
    workbook = xlsxwriter.Workbook(output)

    # Set font
    workbook.formats[0].set_font_name(DEFAULT_FONT)
    workbook.formats[0].set_font_size(DEFAULT_SIZE)

    if metadata:
        write_metadata(args, workbook, field_dict, metadata_df)
    # Create sheet for data
    data_sheet = workbook.add_worksheet('Data')
    variable_sheet_obj = Variable_sheet(workbook)

    header_format = workbook.add_format({
        #         'bg_color': '#C6EFCE',
        'font_color': '#FF0000',
        'font_name': DEFAULT_FONT,
        'bold': False,
        'text_wrap': False,
        'valign': 'vcenter',
        #         'indent': 1,
        'font_size': DEFAULT_SIZE + 2
    })

    field_format = workbook.add_format({
        'font_name': DEFAULT_FONT,
        'bottom': True,
        'right': True,
        'bold': False,
        'text_wrap': True,
        'valign': 'vcenter',
        'font_size': DEFAULT_SIZE + 1,
        'bg_color': '#B9F6F5'
    })

    date_format = workbook.add_format({
        'font_name': DEFAULT_FONT,
        'bold': False,
        'text_wrap': False,
        'valign': 'vcenter',
        'font_size': DEFAULT_SIZE,
        'num_format': 'dd/mm/yy'
        })

    time_format = workbook.add_format({
        'font_name': DEFAULT_FONT,
        'bold': False,
        'text_wrap': False,
        'valign': 'vcenter',
        'font_size': DEFAULT_SIZE,
        'num_format': 'hh:mm:ss'
        })

    title_row = 1  # starting row
    start_row = title_row + 2
    parameter_row = title_row + 1  # Parameter row, hidden
    end_row = 20000  # ending row

    # Loop over all the variables needed
    for ii in range(len(file_def['fields'])):
        # Get the wanted field object
        field = field_dict[file_def['fields'][ii]]

        # Write title row
        data_sheet.write(title_row, ii, field.disp_name, field_format)

        # Write row below with parameter name
        data_sheet.write(parameter_row, ii, field.name)
        # Write validation
        if field.validation is not None:
            if args.verbose > 0:
                print("Writing validation for", file_def['fields'][ii])

            if field.long_list:
                # We need to add the data to the validation sheet
                # Copying the dict as we need to modify it
                valid_copy = field.validation.copy()

                # Add the validation variable to the hidden sheet
                ref = variable_sheet_obj.add_row(
                    field.name, valid_copy['source'])
                valid_copy.pop('source', None)
                valid_copy['value'] = ref
                valid_copy['input_message'].replace('\n', '\n\r')
                data_sheet.data_validation(first_row=start_row,
                                           first_col=ii,
                                           last_row=end_row,
                                           last_col=ii,
                                           options=valid_copy)
            else:
                # Need to make sure that 'input_message' is not more than 255
                valid_copy = field.validation.copy()
                if len(valid_copy['input_message']) > 255:
                    valid_copy['input_message'] = valid_copy[
                        'input_message'][:252] + '...'
                valid_copy['input_message'].replace('\n', '\n\r')
                # Need to make sure that 'input_title' is not more than 32
                # characters long
                if len(valid_copy['input_title']) > 32:
                    valid_copy['input_title'] = valid_copy[
                        'input_title'][:32]

                data_sheet.data_validation(first_row=start_row,
                                           first_col=ii,
                                           last_row=end_row,
                                           last_col=ii,
                                           options=valid_copy)
        if field.cell_format is not None:
            if not('font_name' in field.cell_format):
                field.cell_format['font_name'] = DEFAULT_FONT
            if not('font_size' in field.cell_format):
                field.cell_format['font_size'] = DEFAULT_SIZE
            cell_format = workbook.add_format(field.cell_format)
            data_sheet.set_column(
                ii, ii, width=field.width, cell_format=cell_format)
        else:
            data_sheet.set_column(first_col=ii, last_col=ii, width=field.width)


    # Write optional data to data sheet
    if type(data) == pd.core.frame.DataFrame:
        for col_num, field in enumerate(data):
            try:
                if field in ['eventDate', 'start_date', 'end_date']:
                    data_sheet.write_column(start_row,col_num,list(data[field]), date_format)
                elif field in ['eventTime', 'start_time', 'end_time']:
                    data_sheet.write_column(start_row,col_num,list(data[field]), time_format)
                else:
                    data_sheet.write_column(start_row,col_num,list(data[field]))
            except:
                pass

    # Add header, done after the other to get correct format
    data_sheet.write(0, 0, file_def['disp_name'], header_format)
    # Add hint about pasting
    data_sheet.merge_range(0, 1, 0, 7,
                           "When pasting only use 'paste special' / 'paste only', selecting numbers and/or text ",
                           header_format)
    # Set height of row
    data_sheet.set_row(0, height=24)

    # Freeze the rows at the top
    data_sheet.freeze_panes(start_row, 0)

    # Hide ID row
    data_sheet.set_row(parameter_row, None, None, {'hidden': True})

    if conversions:
        write_conversion(args, workbook)

    workbook.close()


def write_file(url, fields, field_dict, metadata=True, conversions=True, data=False, metadata_df=False):
    """
    Method for calling from other python programs

    Parameters
    ----------
    url: string
        The output file

    fields : list
        A list of the wanted fields

    fields_dict: dict
        A list of the wanted fields on the format shown in config.fields

    metadata: Boolean
        Should the metadata sheet be written
        Default: True

    conversions: Boolean
        Should the conversions sheet be written
        Default: True

    data: pandas.core.frame.DataFrame
        Optional parameter. Option to add data from a dataframe to the 'data' sheet.
        Default: Not added

    metadata_df: pandas.core.frame.DataFrame
        Optional parameter. Option to add metadata from a dataframe to the 'metadata' sheet.
        Default: Not added

    """
    args = Namespace()
    args.verbose = 0
    args.dir = os.path.dirname(url)
    file_def = {'name': os.path.basename(url).split('.')[0],
                'disp_name': '',
                'fields': fields}

    make_xlsx(args, file_def, field_dict, metadata, conversions, data, metadata_df)
