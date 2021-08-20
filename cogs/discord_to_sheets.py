from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import json
import os
import sys
import discord
from discord.ext import commands


class DiscordToSheets(commands.Cog, name="sheet"):
    def __init__(self, bot):
        self.bot = bot
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        self.creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                self.creds = pickle.load(token)

        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'cogs/credentials.json', SCOPES)
                self.creds = flow.run_local_server()
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(self.creds, token)

        self.service = build('sheets', 'v4', credentials=self.creds)
        self.sheet = self.service.spreadsheets().values()

        if not os.path.isfile("config.json"):
            sys.exit("'config.json' not found! Add it and try again.")
        else:
            with open("config.json") as file:
                config = json.load(file)
                self.sheets_id = config["sheets_id"]
                file.close()

#     # set's the bot's google sheet link
#     @commands.command(name="resetsheet",
#                       description="Adds a new Google Sheet (Syntax: `!resetsheet (new sheet's link)`)")
#     @commands.has_role("Active Player")
#     async def reset_sheet(self, context, link):
#         id = link.strip().split("/")[3]
#         if not os.path.isfile("config.json"):
#             sys.exit("'config.json' not found! Add it and try again.")
#         else:
#             with open("config.json") as file:
#                 config = json.load(file)
#                 file.close()
#         config["sheets_id"] = id
#         with open("config.json", "w") as file:
#             json.dump(config, file)
#             file.close()
#         self.sheets_id = id
#         embed = discord.Embed(description="**`Sheet updated!`**", color=0x0a8a14)
#         await context.send(embed=embed)

    # records data on sheet
        @commands.command(name="record", description="Records the provided information on Google Sheets (Syntax: !record "
                                                 "{data})")
    @commands.has_role("Active Player")
    async def record(self, context: commands.Context, content_to_write):

        row_users = self.sheet.get(spreadsheetId=self.sheets_id, range="1:1").execute()  # ['values'][0]
        if 'values' in row_users:
            row_users = row_users['values'][0]
        else:
            row_users = []

        if row_users:
            if str(context.message.author) not in row_users:
                range_column_id = convert_num_to_letters(len(row_users))
                self.write_to_sheet(range=f"{range_column_id}:{range_column_id}", content=str(context.message.author),
                                    append=True)
            else:
                range_column_id = index_to_range(row_users.index(str(context.message.author).strip()) + 1)
        else:
            range_column_id = "B"
            self.write_to_sheet(range="B1", content=str(context.message.author))

        column_days = self.sheet.get(spreadsheetId=self.sheets_id, range="A:A").execute()
        if 'values' in column_days:
            column_days = flatten(column_days['values'])
        else:
            column_days = []

        if column_days:
            msg_date = str(context.message.created_at).split()[0].strip()
            if msg_date not in column_days:
                range_row_id = len(column_days) + 2
                self.write_to_sheet(range=f"{range_row_id}:{range_row_id}", content=msg_date, append=True)
            else:
                range_row_id = column_days.index(msg_date) + 2
        else:
            range_row_id = "2"
            msg_date = str(context.message.created_at).split()[0].strip()
            self.write_to_sheet(range="A2", content=msg_date)

        self.write_to_sheet(range=f"{range_column_id}{range_row_id}", content=content_to_write)

        await context.send(
            embed=discord.Embed(description=f"`{context.message.author}`'s data has been successfully recorded!",
                                color=0x0a8a14))

    def write_to_sheet(self, range, content, append=False):
        if append:
            body = {
                'values': [[content]]
            }
            self.sheet.append(spreadsheetId=self.sheets_id, range=range,
                              valueInputOption='USER_ENTERED', body=body).execute()
        else:
            body = {
                'values': [[content]]
            }
            self.sheet.update(spreadsheetId=self.sheets_id, range=range,
                              valueInputOption='USER_ENTERED', body=body).execute()


# def convert_num_to_letters(last_index):
#     f = lambda x: sys.exit("Can't convert 0 to a letter") if x == 0 else f((x - 1) // 26) + chr((x - 1) % 26 + ord("A"))
#     return f(last_index)


def index_to_range(index):
    return chr(64 + index)


def flatten(t):
    return [item for sublist in t for item in sublist]

def convert_num_to_letters(last_index):
    if last_index:
        index = last_index % 26
        count = (last_index // 26)
        if not index:
            index = 26
        if last_index % 26 == 0:
            count -= 1
        if count:
            range_id = f"{index_to_range(count)}{index_to_range(index)}"
        else:
            range_id = index_to_range(index)
        return range_id
    else:
        sys.exit("Can't convert 0 to a letter")

def setup(bot):
    bot.add_cog(DiscordToSheets(bot=bot))


