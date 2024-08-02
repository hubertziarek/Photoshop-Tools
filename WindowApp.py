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

        label = QLabel(self.directory_name.absolutePath())
        main_layout.addWidget(label)


        doit = QPushButton("Do it")
        doit.clicked.connect(self.add_images_as_layers)
        main_layout.addWidget(doit)

        choose_the_file = QPushButton("Choose the PS file")
        choose_the_file.clicked.connect(self.select_PS_file)
        main_layout.addWidget(choose_the_file)
        
        run_PS = QPushButton("Run PS")
        run_PS.clicked.connect(self.run_PS)
        main_layout.addWidget(run_PS)

        close = QPushButton("Close")
        close.pressed.connect(self.close)
        main_layout.addWidget(close)

        self.show()

    #selecting directory
    def doit_clicked(self):
        #dialog = QFileDialog(self)
        #dialog.setFileMode(QFileDialog.Directory)
        #dialog.setDirectory(QDir("D:\\Orders"))
        #if dialog.exec():
        #    self.directory_name = dialog.directory()
        with Session() as ps:
            desc = ps.ActionDescriptor
            desc.putPath(ps.app.charIDToTypeID("null"), "D:\\Orders\\Andy_thorny_brambles\\thorny_brambles_v1\\arch_1.png")
            event_id = ps.app.charIDToTypeID("Plc ")  # `Plc` need one space in here.
            ps.app.executeAction(ps.app.charIDToTypeID("Plc "), desc)

    #selecting PS file to modify
    def select_PS_file(self):
        dialog = QFileDialog(self)
        dialog.setFileMode(QFileDialog.ExistingFile)
        dialog.setNameFilter("*.psd")
        dialog.setDirectory(QDir("D:\\Orders"))
        if dialog.exec():
            self.PS_file_name = dialog.selectedFiles()

    def run_PS(self):
        #print (self.PS_file_name[0])
        psApp = win32com.client.Dispatch("Photoshop.Application")
        #psApp.Open(self.PS_file_name[0])
        psApp.Open("D:\\Orders\\test\\andy_extras.psd")


    def add_images_as_layers (self):
        #trzeba dodać autościeżkę przy pomocy pathliba, na razie idziemy na sztywno
        directory_path = "D:\\Orders\\test\\renders\\"
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
                    #self.addLayerMask(layer)
                    doc.activeLayer = layer
                    self.moveMe(layer)

            
            for g in groups_to_create:
                # Add a new layerSet.
                new_layer_set = doc.layerSets.add()
                # Rename the layerSet.
                new_layer_set.name = g
                #add new layer
                new_layer = new_layer_set.artLayers.add()
                new_layer.name = "fixes"
            
                layers = doc.artLayers
                layers_to_move = list()
                for layer in layers:
                    if layer.name == g or layer.name == g + "_lineart":
                        layers_to_move.append(layer)

                for o in layers_to_move:
                    o.moveToEnd(new_layer_set)


    def experiments(self):
        pass

    def addLayerMask(self, layer):
        print ("I AM")
        with Session() as ps:
            app = ps.app

            print("Active: " + app.activeDocument.activeLayer.name)
            
            descriptor = ps.ActionDescriptor()
            reference = ps.ActionReference()

            ref = ps.ActionReference()
            ref.putEnumerated(app.stringIDToTypeID("layer"), app.stringIDToTypeID("ordinal"), app.stringIDToTypeID("targetEnum"))
            descriptor.putReference(app.stringIDToTypeID("target"),  ref)

            descriptor.putClass( app.stringIDToTypeID( "new" ), app.stringIDToTypeID( "channel" ))
            reference.putEnumerated( app.stringIDToTypeID( "channel" ), app.stringIDToTypeID( "channel" ), app.stringIDToTypeID( "mask" ))
            descriptor.putReference( app.stringIDToTypeID( "at" ), reference )
            descriptor.putEnumerated( app.stringIDToTypeID( "using" ), app.stringIDToTypeID( "userMaskEnabled" ), app.stringIDToTypeID( "revealAll" ))
            
            app.executeAction( app.stringIDToTypeID( "make" ), descriptor, ps.DialogModes.DisplayNoDialogs )

    def moveMe(self, layer):
        with Session() as ps:
            app = ps.app

            print("Active: " + app.activeDocument.activeLayer.name)
            
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



app = QApplication(sys.argv)
w = MainWindow()
app.exec()