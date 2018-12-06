import pandas as pd
from scipy import sparse
from sklearn.cluster import KMeans
from scipy.linalg import norm
import pathlib
import xlsxwriter
from io import BytesIO
import itertools
from collections import defaultdict
from itertools import count

counter = count()

class idCounter:
    def __init__(self):
        self.id = next(counter)
        self.count = 0
    def increase(self):
        self.count += 1

kategorien = [
'15.0710',
'15.0720',
'15.0730',
'15.0740',
'15.0750',
'13.0020',
'15.0630',
'15.0130',
'15.0160',
'15.0320',
'15.0270',
'15.0285',
'15.0300',
'15.0330',
'15.0340',
'16.0010',
'19.0020',
]


def convertLeistung(l):
    try:
        return '{:07.4f}'.format(float(l))
    except:
        return str(l)

class Categorize:

    maxDist = 1
    randomseed = 5

    def __init__(self, filename):
        self.filename=filename
        if '.csv' in filename:
            self.daten = pd.read_csv(filename,sep=';')
        elif '.xls' in filename:
            self.daten = pd.read_excel(
                    filename,
                    converters = {'Leistung':convertLeistung},
                    )
        else:
            raise Exception("Konnte Datei nicht einlesen")


    def doCalc(self, nCluster = 8):
        self.buildKMeansMat()
        self.fitKmeans(nCluster)
        self.buildGroups()


    def getLeistungen(self):
        datenNurTarmed = self.daten[self.daten.Leistungskategorie == 'Tarmed']
        leistungen = np.unique(datenNurTarmed.Leistung)
        return leistungen

    def buildKMeansMat(self):

        daten = self.daten
        leistungen = self.getLeistungen()

        iIndex = []
        jIndex = []
        for i,group in enumerate(daten.groupby('FallDatum')):
            for j,leistung in enumerate(leistungen):
                if leistung in group[1].Leistung.values:
                    iIndex.append(i)
                    jIndex.append(j)
        daten = np.ones(len(iIndex))
        self.kMeansMat = sparse.coo_matrix( (daten, (iIndex,jIndex) ) )

    def fitKmeans(self, nCluster = 8):

        self.km = KMeans(n_clusters = nCluster, 
                init='k-means++',
                max_iter = 100,
                n_init=1,
                random_state=self.randomseed,
                )

        self.fallKategorie = self.km.fit_predict(self.kMeansMat)

    def writePaketeToExcel(self, filename, directory = 'Resultate'):

        fname = pathlib.Path('.') / directory / filename
        fname = fname.with_suffix('.xlsx')


        def setColWidth(sheet):
            for i, w  in enumerate(lens):
                sheet.set_column(i,i,w)

        daten = self.daten

        lens = [
                1+max([len(str(s)) for s in daten[col].values] + [len(col)]) 
                for col in daten.columns
                ]

        writer = pd.ExcelWriter(str(fname), engine='xlsxwriter')
        daten.to_excel(writer, sheet_name='Rohdaten', index=False)
        workbook = writer.book
        sheet1 = writer.sheets['Rohdaten']

        title_format = workbook.add_format()
        title_format.set_bold()
        title_format.set_bottom()

        cell_format1 = workbook.add_format()
        cell_format2 = workbook.add_format()
        cell_format2.set_bg_color('#dddddd')

        sheet_alle = workbook.add_worksheet('AllePakete')
        counter_alle = 1
        cycler_alle = itertools.cycle( (cell_format1, cell_format2) )

        sheets = [workbook.add_worksheet(l) for l in kategorien]
        sheets.append(workbook.add_worksheet('Restgruppe'))
        sheets.append(workbook.add_worksheet('OhneTarmed'))
        counters = [1 for l in kategorien] + 2*[1]
        cycler = [
                    itertools.cycle( (cell_format1, cell_format2) )
                    for i in range(len(kategorien)+2)
                 ] 


        def getKatNum(group):
            tarmedgroup = group[group.Leistungskategorie=='Tarmed']
            leistungen = tarmedgroup.Leistung.values
            if len(leistungen) == 0:
                return len(kategorien)+1

            for i,k in enumerate(kategorien):
                if k in leistungen:
                    return i
            return i+1


        def createKey(group):
            tarmedgroup = group[group.Leistungskategorie=='Tarmed']
            key = ''.join([
                    str(l) for l in
                    sorted(np.unique((tarmedgroup.Leistung.values)))
                    ])
            return key
        
        pakete = defaultdict(lambda : idCounter())
        for falldatum, group in daten.groupby('FallDatum'):
            key = createKey(group)
            pakete[key].increase()
            daten.loc[group.index,'paketID'] = pakete[key].id

        for falldatum, group in daten.groupby('FallDatum'):
            key = createKey(group)
            daten.loc[group.index,'Anzahl'] = pakete[key].count


        for sheet in [sheet_alle] + sheets:
            sheet.write_row(
                    'A1', 
                    daten.columns.values,
                    title_format,
                    )
            setColWidth(sheet)

        daten=daten.sort_values(['Anzahl','Leistungskategorie'], ascending=False)

        for j,distGroup in daten.groupby('paketID',sort=False):
            for k,fallGroup in distGroup.groupby('FallDatum'):
                katnum = getKatNum(fallGroup)
                clr = next(cycler[katnum])
                clr_alle = next(cycler_alle)
                sheet = sheets[katnum]
                currentRow = counters[katnum]
                for colNum,colName in enumerate(distGroup.columns):
                    sheet.write_column(
                            currentRow,
                            colNum,
                            fallGroup[colName],
                            clr
                            )
                    sheet_alle.write_column(
                            counter_alle,
                            colNum,
                            fallGroup[colName],
                            clr
                            )
                break
            counters[katnum] += len(fallGroup)
            counter_alle += len(fallGroup)
        
        workbook.close()

    def buildGroups(self, restGruppe=False):
        daten = self.daten
        mm = self.kMeansMat.todense()
        fallKategorie = self.fallKategorie
        km = self.km

        falldaten = []
        maxKat = fallKategorie.max()+1

        pakete = defaultdict(lambda : idCounter())
        for i,group in enumerate(daten.groupby('FallDatum')):
            k = fallKategorie[i]
            d = norm(km.cluster_centers_[k,:] - mm[i,:])
            if d > self.maxDist and restGruppe:
                k = maxKat
            falldaten.append( {
                'Falldatum' : group[0],
                'Kategorie' : k,
                'Distanz' : d,
                'Leistungen' : group[1].Leistung.values,
                } )

            key = ''.join([
                    str(l) for l in
                    sorted(np.unique((group[1].Leistung.values)))
                    ])
            pakete[key].increase()

            # daten.loc[group[1].index,'Kategorie'] = fallKategorie[i]
            daten.loc[group[1].index,'Kategorie'] = k
            daten.loc[group[1].index,'Distanz'] = d
            daten.loc[group[1].index,'paketID'] = pakete[key].id

        for falldatum, group in daten.groupby('FallDatum'):
            key = ''.join([
                    str(l) for l in
                    sorted(np.unique((group.Leistung.values)))
                    ])
            daten.loc[group.index,'Anzahl'] = pakete[key].count


        self.falldaten = sorted(
                falldaten, 
                key = lambda tup: (tup['Kategorie'], tup['Distanz'])
                )

    def plotKategorie(self):

        colorPalette = sns.color_palette()
        daten = self.daten

        leistungen = sorted(np.unique(daten.Leistung.values))
        x = []
        y = []
        color = []
        counter = 0
        for falldatum in self.falldaten:
            yi = []
            for leistung in falldatum['Leistungen']:
                yi.append(leistungen.index(leistung))
            y.extend(yi)
            x.extend( (counter * np.ones(len(yi))).tolist() )
            colors = len(yi) * [colorPalette[falldatum['Kategorie'] % 8]]
            # color.extend((falldatum['Kategorie'] * np.ones(len(yi))).tolist() )
            color.extend( colors )
            counter += 1

        f,ax = plt.subplots(figsize=(18,10))
        ax.scatter(x,y,c=color,cmap=sns.color_palette())
        ax.set_xlabel('Falldatum')
        ax.set_ylabel('Leistung')

        ax.axhline(len(self.getLeistungen())+0.5, color='k',linestyle=':')
        return f,ax

    def saveResults(self, directory='Resultate'):
        path = pathlib.Path(self.filename)
        f,ax = self.plotKategorie()

        picPath = path.parent / directory / (path.stem + "_Bild")
        f.savefig(str(picPath.with_suffix('.png')))

        csvPath = path.parent  / directory / (path.stem + "_Eingeteilt")
        daten.to_csv( str(csvPath.with_suffix('.csv')),sep=';',index=False)

    def writeExcel(self,fname, directory='Resultate'):

        fname = pathlib.Path('.') / directory / fname
        fname = fname.with_suffix('.xlsx')

        daten = self.daten
        writer = pd.ExcelWriter(str(fname), engine='xlsxwriter')
        daten.to_excel(writer, sheet_name='Rohdaten', index=False)
        workbook = writer.book
        sheet1 = writer.sheets['Rohdaten']

        f,ax = self.plotKategorie()
        imgdata = BytesIO()
        f.savefig(imgdata, format='png')
        imgdata.seek(0)

        sheet2 = workbook.add_worksheet('Bild')
        sheet2.insert_image("A1", "", options= {'image_data': imgdata})
        plt.close(f)


        title_format = workbook.add_format()
        title_format.set_bold()
        title_format.set_bottom()

        cell_format1 = workbook.add_format()
        cell_format2 = workbook.add_format()
        cell_format2.set_bg_color('#dddddd')


        lens = [
                1+max([len(str(s)) for s in daten[col].values] + [len(col)]) 
                for col in daten.columns
                ]

        def setColWidth(sheet):
            for i, w  in enumerate(lens):
                sheet.set_column(i,i,w)
        
        setColWidth(sheet1)

        sheet = workbook.add_worksheet('Pakete_Alle_Kat')
        sheet.write_row(
                'A1', 
                daten.columns.values,
                title_format,
                )
        currentRow = 1
        cycler = itertools.cycle( (cell_format1, cell_format2) )
        for j,distGroup in daten.groupby('paketID'):
            for k,fallGroup in distGroup.groupby('FallDatum'):
                clr = next(cycler)
                for colNum,colName in enumerate(distGroup.columns):
                    sheet.write_column(
                            currentRow,
                            colNum,
                            fallGroup[colName],
                            clr
                            )
                break
            currentRow += len(fallGroup)
        setColWidth(sheet)


        for i,kategorieGroup in enumerate(daten.groupby('Kategorie')):
            sheet = workbook.add_worksheet('Pakete_Kat_{:02d}'.format(i))
            sheet.write_row(
                    'A1', 
                    kategorieGroup[1].columns.values,
                    title_format,
                    )
            currentRow = 1
            cycler = itertools.cycle( (cell_format1, cell_format2) )
            for j,distGroup in kategorieGroup[1].groupby('paketID'):
                for k,fallGroup in distGroup.groupby('FallDatum'):
                    clr = next(cycler)
                    for colNum,colName in enumerate(distGroup.columns):
                        sheet.write_column(
                                currentRow,
                                colNum,
                                fallGroup[colName],
                                clr
                                )
                    break
                currentRow += len(fallGroup)
            setColWidth(sheet)


        workbook.close()

def doCalcs():
    files = [
            './Rohdaten/Q1-Q3_2018_Angio_alle.xlsx',
            './Rohdaten/Q1-Q3_2018_Endo_alle.xlsx',
            './Rohdaten/Q1-Q3_2018_Gastro_alle.xlsx',
            './Rohdaten/Q1-Q3_2018_RangesPneumo_alle.xlsx',
            ]

    resultFilenames = [
            l.replace('./Rohdaten/','').replace('.xlsx','') + '_Eingeteilt'
            for l in files
            ]

    for f,r in zip(files,resultFilenames):
        c = Categorize(f)
        c.doCalc(20)
        c.writeExcel(r)

