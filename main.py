import sys
import json
from datetime import timedelta

from pyot.models import tft
from pyot.utils import loop_run
from pyot.core import Settings
import xlsxwriter


Settings(
    MODEL="TFT",
    DEFAULT_PLATFORM="NA1",
    DEFAULT_REGION="AMERICAS",
    DEFAULT_LOCALE="EN_US",
    PIPELINE=[
        {
            "BACKEND": "pyot.stores.Omnistone",
            "LOG_LEVEL": 30,
        },
        {
            "BACKEND": "pyot.stores.CDragon",
            "LOG_LEVEL": 30,
            "ERROR_HANDLING": {
                404: ("T", []),
                500: ("R", [3])
            }
        },
        {
            "BACKEND": "pyot.stores.RiotAPI",
            "API_KEY": "RGAPI-23b22ba1-6c6c-45f7-abf6-73c69f9f1dbd",
            "RATE_LIMITER": {
                "BACKEND": "pyot.limiters.MemoryLimiter",
                "LIMITING_SHARE": 1,
            },
            "ERROR_HANDLING": {
                400: ("T", []),
                503: ("E", [3,3])
            }
        }
    ]
).activate()


platform_to_region = {
    "NA1": "AMERICAS",
    "LA1": "AMERICAS",
    "LA2": "AMERICAS",
    "BR1": "AMERICAS",
    "EUW1": "EUROPE",
    "EUN1": "EUROPE",
    "TR1": "EUROPE",
    "RU": "ASIA",
    "JP1": "ASIA",
    "KR": "ASIA",
}


async def main():
    try:
        name = sys.argv[1]
        platform = sys.argv[2].upper()
    except IndexError:
        raise ValueError('Missing arguments for execution (example: python -m tft-comps somename na1)')
    outputs = {}
    summoner = await tft.Summoner(name=name, platform=platform).get()
    history = await tft.MatchHistory(puuid=summoner.puuid, region=platform_to_region[platform]).get()
    match = await history[-1].get() # type: tft.Match
    for participant in match.info.participants:
        p_summoner = await participant.summoner.get()
        outputs[p_summoner.name] = participant.dict()
    outputs = dict(sorted(outputs.items(), key=lambda x: x[1]['placement']))
    json.dump(outputs, open('output.json', 'w+'), indent=4)
    workbook = xlsxwriter.Workbook('output.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True, 'bg_color': '#edf2ac', 'border': 1})
    bold.set_align('center')
    bold.set_align('vcenter')
    wrapcenter = workbook.add_format()
    wrapcenter.set_text_wrap()
    wrapcenter.set_align('center')
    wrapcenter.set_align('vcenter')
    # participant | stage | time alive | styles | units 
    worksheet.set_row(0, 20)
    worksheet.set_column('A:A', 3)
    worksheet.write('A1', '#', bold)
    worksheet.set_column('B:B', 20)
    worksheet.write('B1', 'Participant', bold)
    worksheet.set_column('C:C', 10)
    worksheet.write('C1', 'Stage', bold)
    worksheet.set_column('D:D', 10)
    worksheet.write('D1', 'Alive', bold)
    worksheet.set_column('E:E', 30)
    worksheet.write('E1', 'Traits', bold)
    worksheet.set_column('F:F', 30)
    worksheet.write('F1', 'Units', bold)
    row = 2
    for participant, info in outputs.items():
        worksheet.write(f"A{row}", info['placement'], wrapcenter)
        worksheet.write(f"B{row}", participant, wrapcenter)
        worksheet.write(f"C{row}", f"{info['last_round'] // 5}-{info['last_round'] % 5}", wrapcenter)
        worksheet.write(f"D{row}", str(timedelta(seconds=int(info['time_eliminated']))), wrapcenter)
        worksheet.write(f"E{row}", "".join(map(lambda x: f"{x['name']} (x{x['num_units']})\n", info['traits'])), wrapcenter)
        worksheet.write(f"F{row}", "".join(map(lambda x: f"{x['character_id']} (\u2605{x['tier']})\n", info['units'])), wrapcenter)
        row += 1
    workbook.close()


if __name__ == "__main__":
    loop_run(main())
