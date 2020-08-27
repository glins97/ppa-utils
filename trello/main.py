from trello import TrelloClient
import datetime
import requests
import json
client = TrelloClient(
    api_key='1a3689526b753e44621507d58206e10f',
    api_secret='819914b6634858d877d6269aed6bd42ace24cd4bdaacf2fbdcf3dc91745939ca',
    token='43f878dcce697a84f3368b320458e78c216b81e428702f55e61b79f406e8717f',
)

boards = client.list_boards()
# for board in boards:
#     print(board.name, board.id)

verify = ['5f3e6d1588dccb10dc80ad1e', '5f2af9aff85ab1735ea38f2d']
for board in boards:
    if board.id in verify:
        print(f'{board.name}')
        for card in board.visible_cards():
            if 'tps' not in card.name.lower(): continue
            for attachment in card.attachments:
                try:
                    if 'url' in attachment and 'ppa.digital' in attachment["url"]:
                        if not card.due:
                            print(f'\t -  WARNING::{card.name} NO_DUE_DATE')
                            continue
                        date, time = card.due.split('T')
                        year, month, day = date.split('-')
                        hour, minute, second = time.split(':')
                        id = attachment["url"].split('/')[4]
                        
                        db = json.loads(requests.get(f'https://ppa.digital/tps/delivery_date/{id}').text)
                        card_due = f'{int(year)}-{int(month)}-{int(day)} {int(hour)}-{int(minute)}'
                        db_due = f'{db["year"]}-{db["month"]}-{db["day"]} {db["hour"]+3}-{db["minute"]}'
                        
                        if card_due != db_due:
                            print(f'\t -  WARNING::{card.name} #{id} {card_due} {db_due}')
                        else:
                            print(f'\t -  OK::{card.name} #{id}')
                except Exception as e:
                    print(f'#{id}::ERROR::{e}')