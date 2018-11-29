import pandas as pd
from scipy import sparse
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from scipy.linalg import norm

def readData(filename = './Q1-Q3_2018_RangesEndo.csv'):
    daten = pd.read_csv(filename,sep=';')
    return daten

def getLeistungen(daten):
    datenNurTarmed = daten[daten.Leistungskategorie == 'Tarmed']
    leistungen = np.unique(datenNurTarmed.Leistung)
    return leistungen

def buildKMeansMat(daten):

    leistungen = getLeistungen(daten)

    iIndex = []
    jIndex = []
    for i,group in enumerate(daten.groupby('FallDatum')):
        for j,leistung in enumerate(leistungen):
            if leistung in group[1].Leistung.values:
                iIndex.append(i)
                jIndex.append(j)
    daten = np.ones(len(iIndex))
    return sparse.coo_matrix( (daten, (iIndex,jIndex) ) )

def fitKmeans(daten, nCluster = 8, do_pca=False):

    mat = buildKMeansMat(daten)


    km = KMeans(n_clusters = nCluster, 
            init='k-means++',
            max_iter = 100,
            n_init=1,
            )

    pca = PCA(n_components = 0.9, whiten=True)
    if do_pca:
        transformedData = pca.fit_transform(mat.todense())
    else:
        transformedData = mat
    kategorie = km.fit_predict(transformedData)
    return km, kategorie, pca

def plotKategorie(daten, km, fallKategorie, pca=None, maxDist=1):

    mm = buildKMeansMat(daten).todense()
    if not pca is None:
        mm = pca.fit_transform(mm)

    falldaten = []
    maxKat = fallKategorie.max()+1

    for i,group in enumerate(daten.groupby('FallDatum')):
        k = fallKategorie[i]
        d = norm(km.cluster_centers_[k,:] - mm[i,:])
        if d > maxDist:
            k = maxKat
        falldaten.append( {
            'Falldatum' : group[0],
            'Kategorie' : k,
            'Distanz' : d,
            'Leistungen' : group[1].Leistung.values,
            } )

        daten.loc[group[1].index,'Kategorie'] = k
        daten.loc[group[1].index,'Distanz'] = d

    falldaten = sorted(
            falldaten, 
            key = lambda tup: (tup['Kategorie'], tup['Distanz'])
            )

    colorPalette = sns.color_palette()

    leistungen = sorted(np.unique(daten.Leistung.values))
    x = []
    y = []
    color = []
    counter = 0
    for falldatum in falldaten:
        yi = []
        for leistung in falldatum['Leistungen']:
            yi.append(leistungen.index(leistung))
        y.extend(yi)
        x.extend( (counter * np.ones(len(yi))).tolist() )
        colors = len(yi) * [colorPalette[falldatum['Kategorie'] % 8]]
        # color.extend((falldatum['Kategorie'] * np.ones(len(yi))).tolist() )
        color.extend( colors )
        counter += 1

    f,ax = plt.subplots()
    ax.scatter(x,y,c=color,cmap=sns.color_palette())
    ax.set_xlabel('Falldatum')
    ax.set_ylabel('Leistung')

    ax.axhline(len(getLeistungen(daten)), color='k',linestyle=':')
