from PySide6.QtWidgets import QFileDialog, QApplication, QMainWindow, QLabel, QPushButton, QLineEdit, QComboBox, QCheckBox, QRadioButton, QTextEdit, QSlider, QSpinBox, QProgressBar, QTableWidget, QTableWidgetItem, QVBoxLayout, QHBoxLayout, QWidget
from PySide6.QtCore import QDir
from photoshop import Session
from pathlib import Path

import sys
import win32com.client
import os
import re


class MainWindow(QMainWindow):
    directory_name = QDir("D:\\Orders")
    PS_file_name = str()
    
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Inkarnate assets workflow support")

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        label = QLabel("PHOTOSHOP")
        main_layout.addWidget(label)

        load_images_button = QPushButton("Load images")
        load_images_button.clicked.connect(self.add_images_as_layers)
        main_layout.addWidget(load_images_button)

        export_button = QPushButton("Export images")
        export_button.clicked.connect(self.export_JS)
        main_layout.addWidget(export_button)

        close = QPushButton("Close")
        close.pressed.connect(self.close)
        main_layout.addWidget(close)

        self.show()

    #selecting directory
    def choose_directory(self, starting_path):
        dialog = QFileDialog(self)
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setDirectory(QDir(starting_path))
        if dialog.exec():
            print(dialog.directory().path())
            return dialog.directory().path()
        else:
            return None


    def add_images_as_layers (self):
        directory_path = self.choose_directory("D:\\Orders") + "/"
        if directory_path == None:
            return
        
        images_to_import = [f for f in os.listdir(directory_path) if f.rsplit(".", 1)[1] == "png"]
        groups_to_create = [g.rsplit(".", 1)[0] for g in images_to_import if re.search(".*_lineart.png", g) == None]
        
        with Session() as ps:
            doc = ps.active_document
            for image in images_to_import:
                desc = ps.ActionDescriptor
                desc.putPath(ps.app.charIDToTypeID("null"), directory_path + image)
                ps.app.executeAction(ps.app.charIDToTypeID("Plc "), desc)

            #add masks to lineart
            layers = doc.artLayers
            for layer in layers:
                if re.search(".*_lineart", layer.name) != None:
                    doc.activeLayer = layer
                    self.add_layer_mask()

            
            for g in groups_to_create:
                # Add a new layerSet.
                new_layer_set = doc.layerSets.add()
                # Rename the layerSet.
                new_layer_set.name = g
                #add new layer
                new_layer = new_layer_set.artLayers.add()
                new_layer.name = "fixes"
            
                #move layers to a new group
                layers = doc.artLayers
                layers_to_move = list()
                for layer in layers:
                    if re.search(g, layer.name) != None:
                        layers_to_move.append(layer)

                for o in layers_to_move:
                    o.moveToEnd(new_layer_set)

    def add_layer_mask(self):
        with Session() as ps:
            app = ps.app
            
            idMk = app.charIDToTypeID("Mk  ")
            desc2 = ps.ActionDescriptor()
            idNw = app.charIDToTypeID("Nw  ")
            idChnl = app.charIDToTypeID("Chnl")
            desc2.putClass(idNw, idChnl)
            idAt = app.charIDToTypeID("At  ")
            ref1 = ps.ActionReference()
            idChnl = app.charIDToTypeID("Chnl")
            idChnl = app.charIDToTypeID("Chnl")
            idMsk = app.charIDToTypeID("Msk ")
            ref1.putEnumerated(idChnl, idChnl, idMsk)
            desc2.putReference(idAt, ref1)
            idUsng = app.charIDToTypeID("Usng")
            idUsrM = app.charIDToTypeID("UsrM")
            idRvlA = app.charIDToTypeID("RvlA")
            desc2.putEnumerated(idUsng, idUsrM, idRvlA)
            app.executeAction(idMk, desc2, ps.DialogModes.DisplayNoDialogs)

    def hide_all_layers(self, layers):
        for layer in layers:
            layer.visible = False
    
    def export(self):       
        with Session() as ps:
            doc = ps.active_document
            
            directory_path = self.choose_directory(doc.path) + "/"
            if directory_path == None:
                return

            #merging groups into layers
            groups = doc.layerSets
            duplicated_groups = list()
            for group in groups:
                duplicated_groups.append(group)  
                #group.merge()

            for g in duplicated_groups:
                g.merge()
            
            #export options
            options = ps.PNGSaveOptions()
            options.compression = 1
            
            layers = doc.artLayers
            for layer in layers:
                if re.search(".{1}ackground", layer.name) != None:
                    continue

                self.hide_all_layers(layers)
                layer.visible = True
                print(directory_path)
                if not os.path.exists(directory_path):
                    ps.alert("Directory doesn't exist!")
                    return
                image_path = os.path.join(directory_path, f"{layer.name}.png")
                doc.saveAs(image_path, options=options, asCopy=True)

    def export_JS(self):
        with Session() as ps:
            doc = ps.active_document
            
            directory_path = self.choose_directory(doc.path) + "/"
            if directory_path == None:
                return
            
            layers = doc.artLayers
            for layer in layers:
                if re.search(".{1}ackground", layer.name) != None:
                    continue

                #print(directory_path)
                if not os.path.exists(directory_path):
                    ps.alert("Directory doesn't exist!")
                    return
                
                doc.activeLayer = layer
                self.export_JS_layer(directory_path)

    def export_JS_layer(self, path):
        with Session() as ps:
            app = ps.app

            print (ps.active_document.activeLayer.name)

            d = ps.ActionDescriptor()
            r = ps.ActionReference()

            r.putEnumerated(app.stringIDToTypeID("layer"), app.stringIDToTypeID("ordinal"), app.stringIDToTypeID("targetEnum"))
            d.putReference(app.stringIDToTypeID("null"), r)
            d.putString(app.stringIDToTypeID("fileType"), "png")
            d.putInteger(app.stringIDToTypeID("quality"), 32)
            d.putInteger(app.stringIDToTypeID("metadata"), 0)
            d.putString(app.stringIDToTypeID("destFolder"), path)
            d.putBoolean(app.stringIDToTypeID("sRGB"), True)
            d.putBoolean(app.stringIDToTypeID("openWindow"), False)

            app.executeAction(app.stringIDToTypeID("exportSelectionAsFileTypePressed"), d, ps.DialogModes.DisplayNoDialogs)
            #app.executeAction(app.stringIDToTypeID("exportDocumentAsFileTypePressed"), d, ps.DialogModes.DisplayNoDialogs)
            

app = QApplication(sys.argv)
w = MainWindow()
app.exec()