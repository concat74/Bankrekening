import xml.etree.ElementTree as et
import numpy as np
import pandas as pd
import matplotlib.pyplot as plot
import sys


# Load file into dataframe
def laadBestand(bankbestand):
    df = pd.read_csv( bankbestand, sep=';', na_values=[""] )
    return df

def splitsAfenBij(df, value):
    data = df.loc[df["Af Bij"] == value]
    return data

def maandsheet(df, value):
    data = df.loc[df['Maand'] == value]
    maanddata = XMLread(data)
    return maanddata

def maandinkomsten(df, value):
    data = df.loc[df['Maand'] == value]
    inkomst = data.loc[df["Af Bij"] == 'Bij']
    return inkomst

def maanduitgaven(df, value):
    data = df.loc[df['Maand'] == value]
    uitgaaf = data.loc[df["Af Bij"] == 'Af']
    return uitgaaf

def XMLread(df):
    xtree = et.parse( "/Users/coenvandermaade/Documents/Onderwerpen.xml" )
    xroot = xtree.getroot()
    totalendf = pd.DataFrame( columns=['categorie', 'amount'] )
    for node in xroot:
        s_categorie = node.attrib.get( 'categorie' )
        nodes = node.findall( "name" )
        searchString = ""
        for node in nodes:
            nodedata = df[(df['Naam / Omschrijving'].str.contains( node.text, case=False )) & (df['Af Bij'] == 'Af')]
            nodetotal = nodedata['Bedrag (EUR)'].sum()
            totalendf = pd.concat( [totalendf, pd.DataFrame.from_records( [{'Omschrijving': node.text, 'Bedrag': nodetotal}] )],
                                   ignore_index=True )
            if node == nodes[-1]:
                searchString += node.text
            else:
                searchString += node.text + "|"
            # Zet kruisje dat regel is afgehandeld
            huidigeregel = (df['Naam / Omschrijving'].str.contains( node.text, case=False )) & (df['Af Bij'] == 'Af')
            df.loc[huidigeregel, 'Tag'] = 'x'
        else:
            data = df[(df['Naam / Omschrijving'].str.contains( searchString, case=False )) & (df['Af Bij'] == 'Af')]
            total = data['Bedrag (EUR)'].sum()
            totalendf = pd.concat(
                [totalendf, pd.DataFrame.from_records( [{'Categorie': s_categorie, 'Totaal': total}] )],
                ignore_index=True)
            totalendf = totalendf[['Categorie', 'Omschrijving', 'Bedrag', 'Totaal']]
    return (totalendf)

# Test create different Category sheets
def createsheet(writer, dataframe):
    workbook = writer.book
    bold = workbook.add_format(
        {'bold': True, 'text_wrap': False, "font_size": '14' , 'valign': 'top', 'align': 'center', 'fg_color': '#D7E4BC', 'border': 1} )
    currency = workbook.add_format( {'num_format': '€#,##0.00', 'border':1} )
    for x in range(1,13):
        if x == 1: namesheet = 'Januari'
        if x == 2: namesheet = 'Februari'
        if x == 3: namesheet = 'Maart'
        if x == 4: namesheet = 'April'
        if x == 5: namesheet = 'Mei'
        if x == 6: namesheet = 'Juni'
        if x == 7: namesheet = 'Juli'
        if x == 8: namesheet = 'Augustus'
        if x == 9: namesheet = 'September'
        if x == 10: namesheet = 'Oktober'
        if x == 11: namesheet = 'November'
        if x == 12: namesheet = 'December'
        maand = maandsheet( dataframe, x )
        maand_inkomsten = maandinkomsten( dataframe, x )
        maand_inkomstensum = maand_inkomsten['Bedrag (EUR)'].sum()
        maand_uitgaven = maanduitgaven(dataframe, x)
        maand_uitgavensum = maand_uitgaven['Bedrag (EUR)'].sum()
        maand.to_excel( writer, sheet_name=namesheet, startrow=5, startcol=1 )
        sheet = writer.sheets[namesheet]
        sheet.merge_range( 'B5:F5', namesheet, bold)
        sheet.set_column( 1, 1, 5 )
        sheet.set_column( 2, 5, 25, currency )
        sheet.set_column( 4, 5, 15, currency )
        sheet.write( 'C1', 'totale uitgaven', bold )
        sheet.write( 'C2', maand_uitgavensum, currency )
        sheet.write( 'D1', 'totale inkomsten', bold )
        sheet.write( 'D2', maand_inkomstensum, currency )
        sheet.write( 'E1', 'Verschil', bold )

        # Create a new chart object.
        chart = workbook.add_chart( {'type': 'line'} )
        # Add a series to the chart.
        chart.add_series( {'values': '=Januari!$F$26,=Februari!$F$26,=Maart!$F$26'} )
        # Insert the chart into the worksheet.
        sheet.insert_chart( 'J1', chart )

    return (sheet)

def createInkomstensheet(writer, maanddf):
    inkomsten = maanddf
    inkomsten.to_excel( writer, sheet_name="Inkomsten", startrow=2, startcol=1 )
    inkomstensheet = writer.sheets['Inkomsten']
    inkomstensheet.write( 'B1', "Inkomsten" )
    return (inkomstensheet)


def createUitgavensheet(writer, maanddf):
    uitgaven = maanddf
    uitgaven.to_excel( writer, sheet_name="Uitgaven", startrow=2, startcol=1 )
    uitgavensheet = writer.sheets['Uitgaven']
    uitgavensheet.write( 'B1', "Uitgaven")
    return (uitgavensheet)

# Write data to new Excelfile and create multiple tabs
def schrijfBestand(dataframe):
    filename = "Bankoverzicht_ING.xlsx"
    totalendf = XMLread( dataframe )
    with pd.ExcelWriter( filename, engine='xlsxwriter' ) as writer:

        createsheet(writer, dataframe)
        inkomsten = splitsAfenBij( dataframe, 'Bij' )
        inkomsten_sum = inkomsten["Bedrag (EUR)"].sum()
        inkomsten.to_excel( writer, sheet_name="Inkomsten", startrow=2, startcol=1 )
        inkomstensheet = writer.sheets['Inkomsten']
        inkomstensheet.write( 'B1', "Inkomsten" )

        uitgaven = splitsAfenBij( dataframe, 'Af' )
        uitgaven_sum = uitgaven["Bedrag (EUR)"].sum()
        uitgaven.to_excel( writer, sheet_name="Uitgaven", startrow=2, startcol=1 )
        uitgavensheet = writer.sheets['Uitgaven']
        uitgavensheet.write( 'B1', "Uitgaven" )

        vroegstedatum = uitgaven['Datum'].min()
        laatstedatum = uitgaven['Datum'].max()
        name = "Totaal"

        totalendf.to_excel( writer, sheet_name=name, startrow=4, startcol=1 )
        categoriesheet = writer.sheets[name]
        categoriesheet.set_column( 1, 1, 5 )
        categoriesheet.set_column( 2, 3, 25 )
        categoriesheet.set_column( 4, 5, 15 )
        categoriesheet.write( 'B1', 'Financieel overzicht van' )
        categoriesheet.write( 'C1', vroegstedatum )
        categoriesheet.write( 'D1', laatstedatum )
        categoriesheet.write( 'H1', 'totale inkomsten' )
        categoriesheet.write( 'H2', inkomsten_sum )
        categoriesheet.write( 'I1', 'totale uitgaven' )
        categoriesheet.write( 'I2', uitgaven_sum )


def main():
    bankbestand = sys.argv[1]
    DataFrame = laadBestand(bankbestand)
    DataFrame['Bedrag (EUR)'] = DataFrame['Bedrag (EUR)'].str.replace( ',', '.' ).astype( np.float64 )
    DataFrame['Datum'] = pd.to_datetime( DataFrame['Datum'], format="%Y%m%d", infer_datetime_format=False,
                                         errors='coerce' ).dt.strftime( '%Y-%m-%d' )
    DataFrame['Maand'] = pd.DatetimeIndex( DataFrame['Datum'] ).month
    DataFrame['Tag'] = DataFrame['Tag'].astype( 'string' )
    schrijfBestand(DataFrame )


if __name__ == '__main__':
    main()
