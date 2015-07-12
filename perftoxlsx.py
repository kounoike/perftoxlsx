#!python 
# coding: utf-8

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_range, xl_col_to_name
import click
import datetime
import dateutil.parser
import re
import os.path

def to_datetime(s):
    dt = dateutil.parser.parse(s)
    return dt.strftime("%Y%m%d-%H%M%S")


def write_datasheet(wb, df):
    ws = wb.add_worksheet("data")
    ws.write_column("A2", [to_datetime(x) for x in df.index])
    ws.write_row("B1", df.columns)
    col = 1
    for col, val in df.iteritems():
        ws.write_column(1, col, val.values)
        col += 1


@click.command()
@click.argument("csv", type=click.Path(exists=True))
def to_xlsx(csv):
    """Convert typeperf's CSV to xlsx charts."""
    if not csv.endswith('.csv'):
        click.echo("please select typeperf's csv")
        return
    xlsx = re.sub(r"\.csv$", ".xlsx", csv)
    if os.path.exists(xlsx):
        click.echo("please remove {0}".format(xlsx))
        return
    
    click.echo("convert {0} to {1}".format(csv, xlsx))
    
    df = pd.read_csv(csv, index_col=[0], parse_dates=[0], na_values=[" "])
    df.index.name = None
    writer = pd.ExcelWriter(xlsx, engine="xlsxwriter", datetime_format="yyyy/MM/dd-hh:mm:ss")
    df.to_excel(writer, sheet_name="data")

    workbook = writer.book
    
    for idx in range(df.shape[1]):
        name = "Chart" + xl_col_to_name(idx + 1)
        chartname = df.columns[idx]
        print "idx:{0} name:{1} df:{2}".format(idx, name, chartname)
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
        chart.set_legend({'none': True})
        chart.set_title({'name': chartname, 'overlay': True, 'none': False})
        categories = '=data!' + xl_range(1, 0, df.shape[0], 0)
        values = '=data!' + xl_range(1, idx + 1, df.shape[0], idx + 1)
        chart.add_series({'categories': categories, 'values': values})
        cs = workbook.add_chartsheet(name)
        cs.set_chart(chart)
    workbook.close()


if __name__ == '__main__':
    to_xlsx()