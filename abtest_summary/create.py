import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

class GoogleSheetABTest:
    """
    A class to manage the creation and formatting of a Google Spreadsheet containing AB test results.
    """
    def __init__(self, spreadsheet_id: str, service_account_file: str):

        self.service_account_file = service_account_file
        self.spreadsheet_id = spreadsheet_id
        
        # Colors and formatting
        self.header_border_color = {
            "red": 0.984, 
            "green": 0.737, 
            "blue": 0.015
        }
        self.positive_back = {
            "red": 0.718, 
            "green": 0.882, 
            "blue": 0.804
        }
        self.positive_text = {
            "red": 0.0, 
            "green": 0.627, 
            "blue": 0.51
        }
        self.negative_back = {
            "red": 0.957, 
            "green": 0.78, 
            "blue": 0.765
        }
        self.negative_text = {
            "red": 0.918, 
            "green": 0.263, 
            "blue": 0.208
        }
        self.mid_back = {
            "red": 0.988, 
            "green": 0.910, 
            "blue": 0.698
        }
        self.mid_text = {
            "red": 0.984, 
            "green": 0.737, 
            "blue": 0.015
        }

        self.credentials = Credentials.from_service_account_file(
            self.service_account_file, 
            scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        self.gc = gspread.authorize(self.credentials)
        self.service = build('sheets', 'v4', credentials=self.credentials)

    def _calculate_column_widths(self, df, header_only_cols=range(4, 13), padding=5):
        """Calculates optimal column widths based on content."""
        widths = []
        for i, col in enumerate(df.columns):
            if i in header_only_cols:
                max_len = len(str(col))
            else:
                max_len = max([len(str(col))] + [len(str(v)) for v in df[col]])
            widths.append(max_len + padding)
        return widths

    def _generate_column_width_requests(self, sheet_id, widths):
        """Generates API requests to set column widths."""
        requests = []
        for i, char_len in enumerate(widths):
            pixel_size = int(char_len * 7.2)
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": i,
                        "endIndex": i + 1
                    },
                    "properties": {
                        "pixelSize": pixel_size
                    },
                    "fields": "pixelSize"
                }
            })
        return requests
    
    def _get_white_borders_body_request(self, sheet_id):
        return {
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 5,
                    "endRowIndex": 1000,
                    "startColumnIndex": 0,
                    "endColumnIndex": 26
                },
                "top": {
                    "style": "SOLID",
                    "color": {"red": 1, "green": 1, "blue": 1}
                },
                "bottom": {
                    "style": "SOLID",
                    "color": {"red": 1, "green": 1, "blue": 1}
                },
                "left": {
                    "style": "SOLID",
                    "color": {"red": 1, "green": 1, "blue": 1}
                },
                "right": {
                    "style": "SOLID",
                    "color": {"red": 1, "green": 1, "blue": 1}
                },
                "innerHorizontal": {
                    "style": "SOLID",
                    "color": {"red": 1, "green": 1, "blue": 1}
                },
                "innerVertical": {
                    "style": "SOLID",
                    "color": {"red": 1, "green": 1, "blue": 1}
                }
            }
        }
        
    def _get_font_request(self, sheet_id):
        return {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0, 
                    "endRowIndex": 1000, 
                    "startColumnIndex": 0, 
                    "endColumnIndex": 26
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "fontFamily": "Montserrat"
                        }
                    }
                },
                "fields": "userEnteredFormat.textFormat.fontFamily"
            }
        }
    
    def _get_header_formatting_request(self, sheet_id):
        return {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id, 
                    "startRowIndex": 4, 
                    "endRowIndex": 5
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": self.header_border_color,
                        "horizontalAlignment": "CENTER",
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1.0, 
                                "green": 1.0, 
                                "blue": 1.0
                            },
                            "fontFamily": "Montserrat",
                            "bold": True
                        }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
            }
        }
        
    def _get_header_values_request(self, sheet_id):
        return {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id, 
                    "startRowIndex": 4, 
                    "endRowIndex": 5,
                    "startColumnIndex": 4, 
                    "endColumnIndex": 13
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.0, 
                            "green": 0.627, 
                            "blue": 0.51
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1.0, 
                                "green": 1.0, 
                                "blue": 1.0
                            },
                            "fontFamily": "Montserrat", 
                            "bold": True
                        },
                        "horizontalAlignment": "CENTER"
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
            }
        }

    def _get_header_borders_request(self, sheet_id, color):
        return {
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id, 
                    "startRowIndex": 4, 
                    "endRowIndex": 5, 
                    "startColumnIndex": 0, 
                    "endColumnIndex": 26
                },
                "top": {
                    "style": "SOLID", 
                    "color": color
                }, 
                "bottom": {
                    "style": "SOLID", 
                    "color": color
                },
                "left": {
                    "style": "SOLID", 
                    "color": color
                }, 
                "right": {
                    "style": "SOLID", 
                    "color": color
                },
                "innerHorizontal": {
                    "style": "SOLID", 
                    "color": color
                }, 
                "innerVertical": {
                    "style": "SOLID", 
                    "color": color
                }
            }
        }

    def _get_header_values_borders_request(self, sheet_id):
        return {
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id, 
                    "startRowIndex": 4, 
                    "endRowIndex": 5, 
                    "startColumnIndex": 4, 
                    "endColumnIndex": 13
                },
                "top": {
                    "style": "SOLID", 
                    "width": 1, 
                    "color": {
                        "red": 0.0, 
                        "green": 0.627, 
                        "blue": 0.51
                    }
                },
                "bottom": {
                    "style": "SOLID", 
                    "width": 1, 
                    "color": {
                        "red": 0.0, 
                        "green": 0.627, 
                        "blue": 0.51
                    }
                },
                "left": {
                    "style": "SOLID", 
                    "width": 1, 
                    "color": {
                        "red": 0.0, 
                        "green": 0.627, 
                        "blue": 0.51
                    }
                },
                "right": {
                    "style": "SOLID", 
                    "width": 1, 
                    "color": {
                        "red": 0.0, 
                        "green": 0.627, 
                        "blue": 0.51
                    }
                },
                "innerHorizontal": {
                    "style": "SOLID", 
                    "width": 1, 
                    "color": {
                        "red": 0.0, 
                        "green": 0.627, 
                        "blue": 0.51
                    }
                },
                "innerVertical": {
                    "style": "SOLID", 
                    "width": 1, 
                    "color": {
                        "red": 0.0, 
                        "green": 0.627, 
                        "blue": 0.51
                    }
                }
            }
        }

    def _get_row_alternating_colors(self, sheet_id, num_rows):
        requests = []
        for i in range(num_rows):
            color = {
                "red": 0.95, 
                "green": 0.95, 
                "blue": 0.95
            } if i % 2 == 0 else {
                "red": 1, 
                "green": 1, 
                "blue": 1
            }
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id, 
                        "startRowIndex": i + 5,
                        "endRowIndex": i + 6
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": color
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })
        return requests

    def _get_conditional_formatting_requests(self, sheet_id, alpha):
        return [
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{
                            "sheetId": sheet_id, 
                            "startRowIndex": 5, 
                            "startColumnIndex": 7, 
                            "endColumnIndex": 13
                        }],
                        "booleanRule": {
                            "condition": {
                                "type": "NUMBER_GREATER", 
                                "values": [{
                                    "userEnteredValue": "0"
                                }]
                            },
                            "format": {
                                "backgroundColor": self.positive_back
                            }
                        }
                    }, 
                    "index": 0
                }
            },
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{
                            "sheetId": sheet_id, 
                            "startRowIndex": 5,
                            "startColumnIndex": 7, 
                            "endColumnIndex": 13
                        }],
                        "booleanRule": {
                            "condition": {
                                "type": "NUMBER_LESS", 
                                "values": [{
                                    "userEnteredValue": "0"
                                }]
                            },
                            "format": {
                                "backgroundColor": self.negative_back
                            }
                        }
                    }, 
                    "index": 0
                }
            },
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{
                            "sheetId": sheet_id, 
                            "startRowIndex": 5, 
                            "startColumnIndex": 6, 
                            "endColumnIndex": 7
                        }],
                        "booleanRule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA", 
                                "values": [{
                                    "userEnteredValue": f"=G6>{2 * alpha}"
                                }]
                            },
                            "format": {
                                "backgroundColor": self.negative_back
                            }
                        }
                    }, 
                    "index": 0
                }
            },
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{
                            "sheetId": sheet_id, 
                            "startRowIndex": 5, 
                            "startColumnIndex": 6, 
                            "endColumnIndex": 7
                        }],
                        "booleanRule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA", 
                                "values": [{
                                    "userEnteredValue": f"=G6<{alpha}"
                                }]
                            },
                            "format": {
                                "backgroundColor": self.positive_back
                            }
                        }
                    }, 
                    "index": 0
                }
            },
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{
                            "sheetId": sheet_id, 
                            "startRowIndex": 5, 
                            "startColumnIndex": 6, 
                            "endColumnIndex": 7
                        }],
                        "booleanRule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA", 
                                "values": [{
                                    "userEnteredValue": f"=AND(G6>={alpha}, G6<={2 * alpha})"
                                }]
                            },
                            "format": {
                                "backgroundColor": self.mid_back
                            }
                        }
                    }, 
                    "index": 0
                }
            }
        ]

    def _get_alignment_requests(self, sheet_id, num_columns):
        return [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id, 
                        "startRowIndex": 4, 
                        "endRowIndex": 1000, 
                        "startColumnIndex": 4, 
                        "endColumnIndex": num_columns
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER"
                        }
                    }, 
                    "fields": "userEnteredFormat.horizontalAlignment"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id, 
                        "startRowIndex": 4, 
                        "endRowIndex": 1000, 
                        "startColumnIndex": 0, 
                        "endColumnIndex": 4
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "LEFT"
                        }
                    }, 
                    "fields": "userEnteredFormat.horizontalAlignment"
                }
            }
        ]
        
    def _get_number_format_request(self, sheet_id):
        return {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id, 
                    "startRowIndex": 4, 
                    "endRowIndex": 1000, 
                    "startColumnIndex": 4, 
                    "endColumnIndex": 10
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "NUMBER", 
                            "pattern": "#,##0.00"
                        }, 
                        "horizontalAlignment": "CENTER"
                    }
                },
                "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
            }
        }
        
    def _get_percent_format_request(self, sheet_id):
        return {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id, 
                    "startRowIndex": 4, 
                    "endRowIndex": 1000, 
                    "startColumnIndex": 10, 
                    "endColumnIndex": 13
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "PERCENT", 
                            "pattern": "0.00%"
                        }, 
                        "horizontalAlignment": "CENTER"
                    }
                },
                "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
            }
        }

    def _get_basic_filter_request(self, sheet_id, num_columns, num_rows, start_row_index):
        """Generates the API request to set a basic filter on the data table."""
        return {
            "setBasicFilter": {
                "filter": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_row_index,
                        "endRowIndex": start_row_index + num_rows + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": num_columns
                    }
                }
            }
        }

    def create_summary_sheet(self, df, experiment_name, variant_mapping: dict=None):
        """
        Creates and formats a new spreadsheet with the provided data.
        """
        experiment_name = experiment_name + '_summary'
        
        alpha = df.iloc[0, :]['alpha']
        analysis_type = df.iloc[0, :]['analysis_type'].upper()
        
        df_new = df.copy()
        df_new['%_lift'] = df_new.ate / df_new.control_variant_mean
        df_new['%_ci_lower'] = df_new.ate_ci_lower / df_new.control_variant_mean
        df_new['%_ci_upper'] = df_new.ate_ci_upper / df_new.control_variant_mean
        
        df_new = df_new[[
            'metric_alias', 'treatment_variant_name', 'dimension_name', 'dimension_value',
            'control_variant_mean', 'treatment_variant_mean', 'p_value',
            'ate', 'ate_ci_lower', 'ate_ci_upper',
            '%_lift', '%_ci_lower', '%_ci_upper']]
        
        if variant_mapping:
            df_new['treatment_variant_name'] = df_new['treatment_variant_name'].replace(variant_mapping)
        
        df_new.rename(
            columns={
                'metric_alias': 'Metric', 
                'treatment_variant_name': 'Treatment', 
                'dimension_name': 'Split', 
                'dimension_value': 'Split Value',
                'control_variant_mean': 'Control Mean', 
                'treatment_variant_mean': 'Treatment Mean', 
                'p_value': 'P-Value',
                'ate': 'ATE', 
                'ate_ci_lower': 'ATE Lower', 
                'ate_ci_upper': 'ATE Upper', 
                '%_lift': '%Lift', 
                '%_ci_lower': '%Lift Lower', 
                '%_ci_upper': '%Lift Upper'
            },
            inplace=True
        )
        df_new.loc[df_new['Split'] == '__total_dimension', 'Split'] = 'TOTAL'
        df_new.loc[df_new['Split Value'] == 'total', 'Split Value'] = 'TOTAL'

        add_sheet_request = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': experiment_name,
                        'gridProperties': {
                            'rowCount': len(df_new) + 5,
                            'columnCount': df_new.shape[1]
                        }
                    }
                }
            }]
        }
        
        response = self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body=add_sheet_request
        ).execute()

        sheet_info = response['replies'][0]['addSheet']['properties']
        sheet_id = sheet_info['sheetId']
        worksheet = self.gc.open_by_key(self.spreadsheet_id).worksheet(experiment_name)
        
        # Write "Test Configurations" table
        worksheet.update(
            [
                ['Test Configurations', ''], 
                ['Type', analysis_type], 
                ['Alpha', alpha]
            ], 'A1'
        )

        # Paste the data starting from row 5
        worksheet.update(
            [df_new.columns.values.tolist()] + df_new.values.tolist(), 
            'A5'
        )

        # List of formatting requests
        all_requests = []
        
        # Add white borders from row 1 to 4
        all_requests.append({
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 4,
                    "startColumnIndex": 0,
                    "endColumnIndex": 26
                },
                "top": {"style": "SOLID", "color": {"red": 1, "green": 1, "blue": 1}},
                "bottom": {"style": "SOLID", "color": {"red": 1, "green": 1, "blue": 1}},
                "left": {"style": "SOLID", "color": {"red": 1, "green": 1, "blue": 1}},
                "right": {"style": "SOLID", "color": {"red": 1, "green": 1, "blue": 1}},
                "innerHorizontal": {"style": "SOLID", "color": {"red": 1, "green": 1, "blue": 1}},
                "innerVertical": {"style": "SOLID", "color": {"red": 1, "green": 1, "blue": 1}}
            }
        })

        # Add "Test Configurations" table formatting
        config_table_header_range = {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": 1,
            "startColumnIndex": 0,
            "endColumnIndex": 2
        }
        all_requests.append({
            "repeatCell": {
                "range": config_table_header_range,
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": self.header_border_color,
                        "textFormat": {
                            "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                            "fontFamily": "Montserrat",
                            "bold": True
                        }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
            }
        })
        all_requests.append({
            "updateBorders": {
                "range": config_table_header_range,
                "top": {"style": "SOLID", "color": self.header_border_color},
                "bottom": {"style": "SOLID", "color": self.header_border_color},
                "left": {"style": "SOLID", "color": self.header_border_color},
                "right": {"style": "SOLID", "color": self.header_border_color},
                "innerHorizontal": {"style": "SOLID", "color": self.header_border_color},
                "innerVertical": {"style": "SOLID", "color": self.header_border_color}
            }
        })
        # Center align for B1-B3
        all_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 3,
                    "startColumnIndex": 1,
                    "endColumnIndex": 2
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER"
                    }
                },
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })
        
        # Add all formatting requests for the main table
        all_requests.append(self._get_font_request(sheet_id))
        all_requests.append(self._get_white_borders_body_request(sheet_id))
        all_requests.append(self._get_header_borders_request(sheet_id, self.header_border_color))
        all_requests.append(self._get_header_formatting_request(sheet_id))
        all_requests.append(self._get_header_values_request(sheet_id))
        all_requests.append(self._get_header_values_borders_request(sheet_id))
        all_requests.extend(self._get_row_alternating_colors(sheet_id, len(df_new)))
        all_requests.extend(self._get_conditional_formatting_requests(sheet_id, alpha))
        all_requests.append(self._get_alignment_requests(sheet_id, df_new.shape[1]))
        all_requests.append(self._get_number_format_request(sheet_id))
        all_requests.append(self._get_percent_format_request(sheet_id))
        
        widths = self._calculate_column_widths(df_new)
        resize_requests = self._generate_column_width_requests(sheet_id, widths)
        all_requests.extend(resize_requests)

        # Aggiungi la richiesta per i filtri
        num_columns = df_new.shape[1]
        num_rows = len(df_new)
        all_requests.append(self._get_basic_filter_request(sheet_id, num_columns, num_rows, start_row_index=4))

        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={"requests": all_requests}
        ).execute()

        print(f"✅ New Sheet '{experiment_name}' added to https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}")
        