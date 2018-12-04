import pandas as pd
from scipy import sparse
from sklearn.cluster import KMeans
from scipy.linalg import norm
import pathlib
import xlsxwriter
from io import BytesIO
import itertools
from collections import defaultdict

class Node:
    def __init__(self,Leistungen):
        self.Leistungen = Leistungen
        self.count = 1
        self.children = defaultdict(int)

    def append(self, Leistungsset):
        if self.Leistungen == Leistungsset:
            self.count+=1
            return True
        else:
            for child in self.children:
                if child.append(Leistungsset):
                    self.children[child] += 1
                    return True
        if self.Leistungen < Leistungsset:
            self.children[Node(Leistungsset)] = 1
            return True

        return False


class Tree:
    def __init__(self,daten):
        self.mainNode = Node(set())
        self.daten = daten

    def sortIntoNodes(self):
        falldaten = []
        for falldatum,group in daten.groupby('FallDatum'):
            leistungsSet = set(group.Leistung.values)
            falldaten.append(leistungsSet)
        falldaten = sorted(falldaten, key=lambda x : len(x))
        for f in falldaten:
            self.mainNode.append(f)

    def writeToExcel(self):
        pass




    

class Categorize:

    maxDist = 1
    randomseed = 5

    def __init__(self, filename):
        self.filename=filename
        self.daten = pd.read_csv(filename,sep=';')

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

    def buildGroups(self):
        daten = self.daten
        mm = self.kMeansMat.todense()
        fallKategorie = self.fallKategorie
        km = self.km

        falldaten = []
        maxKat = fallKategorie.max()+1

        for i,group in enumerate(daten.groupby('FallDatum')):
            k = fallKategorie[i]
            d = norm(km.cluster_centers_[k,:] - mm[i,:])
            if d > self.maxDist:
                k = maxKat
            falldaten.append( {
                'Falldatum' : group[0],
                'Kategorie' : k,
                'Distanz' : d,
                'Leistungen' : group[1].Leistung.values,
                } )

            # daten.loc[group[1].index,'Kategorie'] = fallKategorie[i]
            daten.loc[group[1].index,'Kategorie'] = k
            daten.loc[group[1].index,'Distanz'] = d

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
        for i,kategorieGroup in enumerate(daten.groupby('Kategorie')):
            sheet = workbook.add_worksheet('Pakete_Kat_{:02d}'.format(i))
            sheet.write_row(
                    'A1', 
                    kategorieGroup[1].columns.values,
                    title_format,
                    )
            currentRow = 1
            cycler = itertools.cycle( (cell_format1, cell_format2) )
            for j,distGroup in kategorieGroup[1].groupby('Distanz'):
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
