from codecs import BOM
from tkinter import PROJECTING
import win32com.client as com
from ctypes import *
import os

mod = com.gencache.EnsureModule('{D98A091D-3A0F-4C3E-B36E-61F62068D488}', 0, 1, 0)

BOM_structure_names = {
    51969: "Defecto",
    51970: "Normal",
    51971: "Fantasma",
    51972: "Referencia",
    51973: "Comprado",
    51974: "Inseparable",
    51975: "Varias estructuras"
}

def BOMRowsToArray(prev_node, project, Rows):
    rows = []
    child_rows = []
    components_paths = []
    for i in range(1, Rows.Count+1):
        lRow = Rows.Item(i)
        row_component = lRow.ComponentDefinitions.Item(1).Document.PropertySets.Item('Design Tracking Properties')
        row_component = mod.PropertySet(row_component)
        node_id = prev_node + str(i) + "/"
        if lRow.ChildRows is not None: child_rows.append((node_id, lRow.ChildRows)) 
        part_number = row_component.Item("Part Number").Value
        BOM_structure = lRow.BOMStructure
        quantity = lRow.ItemQuantity
        component_path =  lRow.ComponentDefinitions.Item(1).Document.FullDocumentName
        components_paths.append((part_number, component_path))
        rows.append((node_id, project, part_number, BOM_structure_names[BOM_structure], quantity))
        print(f"Elemento: ({node_id}, {project}, {part_number}, {BOM_structure_names[BOM_structure]}, {quantity}) con ruta {component_path}")
    if child_rows is not None:
        for child in child_rows:
            x1, x2 = BOMRowsToArray(child[0], project, child[1])
            rows.extend(x1)
            components_paths.extend(x2)
    return [rows, components_paths]

def savePathsArrayToFile(ar):
    file = open("component_paths.csv", "w")
    for element in ar:
        file.write(element[0]+';'+element[1]+'\n')

def uploadDirectory(path, bucket, project):
        for root, dirs, files in os.walk(path):
            for file in files:
                bucket.upload_file(os.path.join(root, file), project + "/" + file)

def verifyProjectEntry(pr, c):
    query = "SELECT COUNT(project_name) FROM projects WHERE project_name='{}'".format(pr)
    c.execute(query)
    if c.fetchone()[0] == 0:
        query = "INSERT INTO projects VALUES ('{}')".format(pr)
        c.execute(query)
