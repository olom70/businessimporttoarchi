from turtle import width
from executing import Source
import pylightxl as xl
import platform
import lib.csvutil as csvutil
import lib.stringutil as stringutil
import os
import time
import pyyed

def main():

    if (platform.system() == 'Linux'):
        MAIN_FOLDER = "~home/importVIPHierarchy"
    else:
        MAIN_FOLDER ="C:/Users/MOREAUCL/Documents/importVIPHierarchy"

    IMPORTED_FILE = 'VIP - Exigences - Offre Outil Cible.xlsx'
    YEDFILE = 'VIP.graphml'
    INPUT = MAIN_FOLDER + os.path.sep + IMPORTED_FILE
    ELEMENTTYPE = 'Grouping'
    RELATIONTYPE = 'CompositionRelationship'
    KEY = 'import'
    VALUE = 'VIP'
    L1FONTSTYLE = "bold"
    L1FONTSIZE = "14"
    L2FONTSTYLE = "plain"
    L2FONTSIZE = "13"
    L3FONTSTYLE = "plain"
    L3FONTSIZE = "12"
    WIDTH='200'
    L1COLOR="#DE1FF2"
    L2COLOR="#07F2B8"
    L3COLOR="#F2BE1F"

    try:
        os.mkdir(MAIN_FOLDER)
    except FileExistsError as fe:
        pass
    except FileNotFoundError as fnf:
        print('wrong path, correct MAIN_FOLDER') 
        exit()
    except Exception as e:
        print(f'unexpected error : {type(e)}{e.args}')
        exit()

    if not os.path.isfile(INPUT):
        print(f'this file does not exists {input}. put it in the folder {MAIN_FOLDER}')

    db = xl.readxl(fn=INPUT)

    creationtime = str(time.time())
    csvelementsfile = MAIN_FOLDER + os.path.sep + 'vip-' + creationtime + '-elements.csv'
    csvpropertiesfile = MAIN_FOLDER + os.path.sep + 'vip-' + creationtime + '-properties.csv'
    csvrelationsfile = MAIN_FOLDER + os.path.sep + 'vip-' + creationtime + '-relations.csv'
    outputfiles = csvutil.createfiles([csvelementsfile, csvpropertiesfile, csvrelationsfile])
    outputfiles[1].writerow(csvutil.initElementsHeader())
    outputfiles[3].writerow(csvutil.initPropertiesHeader())
    outputfiles[5].writerow(csvutil.initRelationsHeader())

    colId =  db.ws(ws='Exigences-besoins').col(col=2)
    colCategory = db.ws(ws='Exigences-besoins').col(col=3)
    colPerimeter = db.ws(ws='Exigences-besoins').col(col=5)
    colthematic = db.ws(ws='Exigences-besoins').col(col=6)
    colgroup = db.ws(ws='Exigences-besoins').col(col=7)
    colUseCase = db.ws(ws='Exigences-besoins').col(col=8)
    colDescription = db.ws(ws='Exigences-besoins').col(col=10)
    colPriority = db.ws(ws='Exigences-besoins').col(col=14)

    l_alreadyAdded = []
    g = pyyed.Graph()
    g.define_custom_property("node", "UseCase", "string", "")

    for items in zip(colId, colCategory, colPerimeter, colthematic, colgroup, colUseCase, colDescription, colPriority):
#                      0       1            2               3          4         5             6              7
        if items[1] == 'xx':  # coulb be used to avoid processing 'Besoins'
            pass
        else:
            # IDgeneration ####################################################
            Level1ID = stringutil.cleanName(
                                                items[2],
                                                True,
                                                True,
                                                'lowercase',
                                                True,
                                                True,
                                                True) \
                            + '_Level1'
            if Level1ID not in ['NoName_Level1', 'périmètre_Level1', 'thématique_Level1', 'rgpt_Level1']:
                Level2ID = Level1ID+stringutil.cleanName(
                                                    items[3],
                                                    True,
                                                    True,
                                                    'lowercase',
                                                    True,
                                                    True,
                                                    True) \
                                + '_Level2'
                Level3ID = Level2ID+stringutil.cleanName(
                                                    items[4],
                                                    True,
                                                    True,
                                                    'lowercase',
                                                    True,
                                                    True,
                                                    True) \
                                + '_Level3'
                
                #Level1 ###############################################################1
                if Level1ID not in l_alreadyAdded:
                    l_alreadyAdded.append(Level1ID)

                    Name = stringutil.cleanName(
                                                        items[2],
                                                        False,
                                                        False,
                                                        'nochange',
                                                        True,
                                                        False,
                                                        False) \
                            + ' (L1)'
                    Documentation = ''
                    toWrite= csvutil.initElements(ID=Level1ID, Type=ELEMENTTYPE, Name=Name, Documentation=Documentation)
                    outputfiles[1].writerow(toWrite)

                    toWrite = csvutil.initProperties(ID=Level1ID, Key=KEY, Value=VALUE)
                    outputfiles[3].writerow(toWrite)

                    g.add_node(Level1ID, label=Name, font_size=L1FONTSIZE, font_style=L1FONTSTYLE, width=WIDTH, shape_fill=L1COLOR)

                
                #Level2 ###############################################################
                if Level2ID not in l_alreadyAdded:
                    l_alreadyAdded.append(Level2ID)
                    Name = stringutil.cleanName(
                                                        items[3],
                                                        False,
                                                        False,
                                                        'nochange',
                                                        True,
                                                        False,
                                                        False) \
                            + ' (L2)'
                    Documentation = ''
                    toWrite= csvutil.initElements(ID=Level2ID, Type=ELEMENTTYPE, Name=Name, Documentation=Documentation)
                    outputfiles[1].writerow(toWrite)

                    toWrite = csvutil.initProperties(ID=Level2ID, Key=KEY, Value=VALUE)
                    outputfiles[3].writerow(toWrite)

                    toWrite = csvutil.initRelations(ID=Level2ID, Type=RELATIONTYPE, Source=Level1ID, Target=Level2ID)
                    outputfiles[5].writerow(toWrite)

                    g.add_node(Level2ID, label=Name, font_size=L2FONTSIZE, font_style=L2FONTSTYLE, width=WIDTH, shape_fill=L2COLOR)
                    g.add_edge(Level1ID, Level2ID)


                #Level 3 ##############################################################
                if Level3ID not in l_alreadyAdded:
                    l_alreadyAdded.append(Level3ID)
                    Name = stringutil.cleanName(
                                                        items[4],
                                                        False,
                                                        False,
                                                        'nochange',
                                                        True,
                                                        False,
                                                        False) \
                            + ' (L3)'
                    Doc1 = stringutil.cleanName(
                                                        items[5],
                                                        False,
                                                        False,
                                                        'nochange',
                                                        True,
                                                        False,
                                                        False)
                    Doc2 = stringutil.cleanName(
                                                        items[6],
                                                        False,
                                                        False,
                                                        'nochange',
                                                        True,
                                                        False,
                                                        False)
                
                    toWrite= csvutil.initElements(ID=Level3ID, Type=ELEMENTTYPE, Name=Name, Documentation=Doc1+Doc2)
                    outputfiles[1].writerow(toWrite)

                    toWrite = csvutil.initProperties(ID=Level3ID, Key=KEY, Value=VALUE)
                    outputfiles[3].writerow(toWrite)

                    toWrite = csvutil.initRelations(ID=Level3ID, Type=RELATIONTYPE, Source=Level2ID, Target=Level3ID)
                    outputfiles[5].writerow(toWrite)

                    g.add_node(Level3ID, label=Name, font_size=L3FONTSIZE, font_style=L3FONTSTYLE, width=WIDTH, shape_fill=L3COLOR,
                               custom_properties={"UseCase": Doc1})
                    g.add_edge(Level2ID, Level3ID)
                    # g.add_node(Level3ID+'B', label=items[5]+items[6], font_size=L3FONTSIZE, font_style=L3FONTSTYLE)
                    # g.add_edge(Level3ID, Level3ID+'B', label="use case is")
            
                    g.write_graph(MAIN_FOLDER + os.path.sep + YEDFILE, pretty_print=True)


if __name__ == '__main__':
    main()
