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

    max_traits = 0
    max_units = 0
    for participant, info in outputs.items():
        traits_len = len(info['traits'])
        units_len = len(info['units'])
        if traits_len > max_traits:
            max_traits = traits_len
        if units_len > max_units:
            max_units = units_len
        
    json.dump(outputs, open('output.json', 'w+'), indent=4)

    workbook = xlsxwriter.Workbook('output.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True, 'bg_color': '#edf2ac', 'border': 1})
    bold.set_align('center')
    bold.set_align('vcenter')
    wrapcenter = workbook.add_format()
    wrapcenter.set_align('center')
    wrapcenter.set_align('vcenter')
    # participant | stage | time alive | styles | units 
    worksheet.set_row(0, 20)
    worksheet.set_column('A:A', 3)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:D', 10)
    worksheet.set_column(4, 3+max_traits+max_units, 25)
    worksheet.write('A1', '#', bold)
    worksheet.write('B1', 'Participant', bold)
    worksheet.write('C1', 'Stage', bold)
    worksheet.write('D1', 'Alive', bold)
    worksheet.merge_range(0, 4, 0, 3+max_traits, 'Traits', bold)
    worksheet.merge_range(0, 4+max_traits, 0, 3+max_traits+max_units, 'Units', bold)
    row = 2
    for participant, info in outputs.items():
        worksheet.write(f"A{row}", info['placement'], wrapcenter)
        worksheet.write(f"B{row}", participant, wrapcenter)
        worksheet.write(f"C{row}", f"{info['last_round'] // 5}-{info['last_round'] % 5}", wrapcenter)
        worksheet.write(f"D{row}", str(timedelta(seconds=int(info['time_eliminated']))), wrapcenter)
        for ind, trait in enumerate(map(lambda x: f"{x['name']} (x{x['num_units']})\n", info['traits'])):
            worksheet.write(row-1, 4+ind, trait, wrapcenter)
        for ind, unit in enumerate(map(lambda x: f"{x['character_id']} (\u2605{x['tier']})\n", info['units'])):
            worksheet.write(row-1, 4+max_traits+ind, unit, wrapcenter)
        row += 1
    workbook.close()


if __name__ == "__main__":
    loop_run(main())
