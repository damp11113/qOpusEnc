import math

from PyQt5.QtWidgets import QApplication, QMainWindow, QAbstractItemView, QTableWidgetItem, \
    QFileDialog, QMenu, QAction, QDialog, QPushButton, QVBoxLayout, QLabel
from PyQt5.QtCore import Qt, pyqtSignal, QObject, QThread
from pathlib import Path
import numpy as np
import subprocess
import importlib
import tempfile
import window
import pyogg
import json
import wave
import time
import os

class opusconfigstorer:
    def __init__(self):
        self.version = 4
        self.framesizes = 60
        self.app = "audio"
        self.bitmode = "CVBR"
        self.bitrate = 64.0
        self.bandwidth = "fullband"
        self.channel = 2
        self.packloss = 0
        self.compression = 10
        self.samplesrates = 48000
        self.ai = False
        self.prediction = True
        self.phaseinvert = True
        self.DTX = False
        self.gain = 1
        self.absthreshold = -75
        self.amthreshold = -50
        self.automono = False
        self.dabs = False
        self.absavg = False
        self.oldcodebook = False

    def save(self, name):
        settings = {
            'version': self.version,
            'framesizes': self.framesizes,
            'app': self.app,
            'bitmode': self.bitmode,
            'bitrates': self.bitrate,
            'bandwidth': self.bandwidth,
            'channels': self.channel,
            'packloss': self.packloss,
            'compressions': self.compression,
            'samplesrates': self.samplesrates,
            'ai': self.ai,
            'prediction': self.prediction,
            'phaseinvert': self.phaseinvert,
            'DTX': self.DTX,
            "gain": self.gain,
            "absthreshold": self.absthreshold,
            "amthreshold": self.amthreshold,
            "automono": self.automono,
            "dabs": self.dabs,
            "absavg": self.absavg,
            "oldcodebook": self.oldcodebook
        }

        with open(f"./presets/{name}.json", 'w') as jsonfile:
            json.dump(settings, jsonfile, indent=4)

    def remove(self, name):
        os.remove(f"./presets/{name}.json")

    def load(self, name, ui):
        with open(f"./presets/{name}.json", 'r') as jsonfile:
            settings = json.load(jsonfile)
            # set version
            self.version = settings.get('version', self.version)
            if self.version == 1:
                ui.Use120ms.setEnabled(True)
                ui.Use100ms.setEnabled(True)
                ui.Use80ms.setEnabled(True)
            else:
                ui.Use120ms.setEnabled(False)
                ui.Use100ms.setEnabled(False)
                ui.Use80ms.setEnabled(False)

            if self.version == 1:
                ui.UseHEV2opus.setChecked(True)
            elif self.version == 2:
                ui.UseHEopus.setChecked(True)
            elif self.version == 3:
                ui.UseNEWopus.setChecked(True)
            elif self.version == 4:
                ui.UseSTABLEopus.setChecked(True)
            elif self.version == 5:
                ui.UseOLDopus.setChecked(True)
            else:
                ui.UseSTABLEopus.setChecked(True)
            # set Frame size
            self.framesizes = settings.get('framesizes', self.framesizes)
            if self.framesizes == 120:
                ui.Use120ms.setChecked(True)
            elif self.framesizes == 100:
                ui.Use100ms.setChecked(True)
            elif self.framesizes == 80:
                ui.Use80ms.setChecked(True)
            elif self.framesizes == 60:
                ui.Use60ms.setChecked(True)
            elif self.framesizes == 40:
                ui.Use40ms.setChecked(True)
            elif self.framesizes == 20:
                ui.Use20ms.setChecked(True)
            elif self.framesizes == 10:
                ui.Use10ms.setChecked(True)
            elif self.framesizes == 5:
                ui.Use5ms.setChecked(True)
            elif self.framesizes == 2.5:
                ui.Use2d5ms.setChecked(True)
            else:
                ui.Use60ms.setChecked(True)
            # set app
            self.app = settings.get('app', self.app)
            if self.app == "restricted_lowdelay":
                ui.UseCELTApp.setChecked(True)
            elif self.app == "audio":
                ui.UseHybridApp.setChecked(True)
            elif self.app == "voip":
                ui.UseSILKApp.setChecked(True)
            else:
                ui.UseCELTApp.setChecked(True)
            # set bitrate mode
            self.bitmode = settings.get('bitmode', self.bitmode)
            if self.bitmode == "VBR":
                ui.UseVBRbm.setChecked(True)
            elif self.bitmode == "CVBR":
                ui.UseCVBRbm.setChecked(True)
            elif self.bitmode == "CBR":
                ui.UseCBRbm.setChecked(True)
            else:
                ui.UseCVBRbm.setChecked(True)
            # set bandwidth
            self.bandwidth = settings.get('bandwidth', self.bandwidth)
            if self.bandwidth == "auto":
                ui.UseAUTObw.setChecked(True)
            elif self.bandwidth == "fullband":
                ui.UseFBbw.setChecked(True)
            elif self.bandwidth == "superwideband":
                ui.UseSWBbw.setChecked(True)
            elif self.bandwidth == "wideband":
                ui.UseWBbw.setChecked(True)
            elif self.bandwidth == "mediumband":
                ui.UseMBbw.setChecked(True)
            elif self.bandwidth == "narrowband":
                ui.UseNBbw.setChecked(True)
            else:
                ui.UseFBbw.setChecked(True)
            # set bitrates value
            self.bitrate = settings.get('bitrates', self.bitrate)
            ui.BitratesIn.setValue(self.bitrate)
            # set channels value
            self.channel = settings.get('channels', self.channel)
            ui.ChannelIn.setValue(self.channel)
            # set packet loss value
            self.packloss = settings.get('packloss', self.packloss)
            ui.PACLOSIn.setValue(self.packloss)
            # set compression complex value
            self.compression = settings.get('compressions', self.compression)
            ui.CompIn.setValue(self.compression)
            # set samples rates value
            self.samplesrates = settings.get('samplesrates', self.samplesrates)
            ui.SampRateIn.setValue(self.samplesrates)
            # set gain
            self.gain = settings.get('gain', self.gain)
            ui.GainIn.setValue(self.gain)
            # set enable ai
            self.ai = settings.get('ai', self.ai)
            ui.EnaABS.setChecked(self.ai)
            # set enable prediction
            self.prediction = settings.get('prediction', self.prediction)
            ui.EnaPred.setChecked(self.prediction)
            # set enable phase invert
            self.phaseinvert = settings.get('phaseinvert', self.phaseinvert)
            ui.EnaSPI.setChecked(self.phaseinvert)
            # set enable DTX
            self.DTX = settings.get('DTX', self.DTX)
            ui.EnaDTX.setChecked(self.DTX)

            # set abs threshold
            self.absthreshold = settings.get('absthreshold', self.absthreshold)
            ui.ABSTIn.setValue(self.absthreshold)
            # set am threshold
            self.amthreshold = settings.get('amthreshold', self.amthreshold)
            ui.AMTIn.setValue(self.amthreshold)

            self.automono = settings.get('automono', self.automono)
            ui.EnaAM.setChecked(self.automono)

            self.dabs = settings.get('dabs', self.dabs)
            ui.EnaDABS.setChecked(self.dabs)

            self.absavg = settings.get('absavg', self.absavg)
            ui.EnaABSAvg.setChecked(self.absavg)

            self.oldcodebook = settings.get('oldcodebook', self.oldcodebook)
            ui.EnaABSOCB.setChecked(self.oldcodebook)


    def getopusstrversion(self):
        if self.version == 1:
            return "hev2"
        elif self.version == 2:
            return "he"
        elif self.version == 3:
            return "exper"
        elif self.version == 4:
            return "stable"
        elif self.version == 5:
            return "old"
        else:
            return "stable"

    def setopusencoder(self, opusencoder):
        opusencoder.set_application(self.app)
        opusencoder.set_sampling_frequency(self.samplesrates)
        opusencoder.set_channels(self.channel)
        opusencoder.set_bitrates(int(self.bitrate * 1000))
        opusencoder.set_bandwidth(self.bandwidth)
        opusencoder.set_compresion_complex(self.compression)
        opusencoder.set_bitrate_mode(self.bitmode)
        opusencoder.set_frame_size(self.framesizes)
        opusencoder.set_packets_loss(self.packloss)

        opusencoder.CTL(pyogg.opus.OPUS_SET_PREDICTION_DISABLED_REQUEST, int(self.prediction))
        opusencoder.CTL(pyogg.opus.OPUS_SET_PHASE_INVERSION_DISABLED_REQUEST, int(self.phaseinvert))
        opusencoder.CTL(pyogg.opus.OPUS_SET_DTX_REQUEST, int(self.DTX))

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super(AboutDialog, self).__init__(parent)

        self.setWindowTitle("About")
        self.setFixedSize(400, 300)

        layout = QVBoxLayout()

        # Information labels
        self.lbl_name = QLabel("qOpus Encoder ")
        self.lbl_name.setStyleSheet("font-weight: bold; font-size: 14pt;")
        self.lbl_version = QLabel("Version: 1.5")
        self.lbl_author = QLabel("damp11113")
        self.lbl_website = QLabel('<a href="https://dpsoftware.damp11113.xyz/qopusenc">https://dpsoftware.damp11113.xyz/qopusenc</a>')
        self.lbl_website.setOpenExternalLinks(True)

        layout.addWidget(self.lbl_name)
        layout.addWidget(self.lbl_version)
        layout.addWidget(self.lbl_author)
        layout.addWidget(self.lbl_website)

        layout.addSpacing(20)

        # Other open source projects labels
        self.lbl_ffmpeg = QLabel('<a href="https://ffmpeg.org">FFmpeg</a>')
        self.lbl_ffmpeg.setOpenExternalLinks(True)
        self.lbl_media_info = QLabel('<a href="https://github.com/damp11113/PyOgg">PyOgg (damp11113 moded)</a>')
        self.lbl_media_info.setOpenExternalLinks(True)

        layout.addWidget(QLabel("FFmpeg"))
        layout.addWidget(self.lbl_ffmpeg)
        layout.addWidget(QLabel("PyOgg"))
        layout.addWidget(self.lbl_media_info)

        layout.addSpacing(20)

        # Other open source projects used in the program
        other_projects_label = QLabel("Other open source projects used in the program")
        other_projects_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(other_projects_label)

        other_projects = [
            "PyQt: https://www.riverbankcomputing.com/software/pyqt/"
            # Add more projects here if needed
        ]

        for project in other_projects:
            lbl_project = QLabel(project)
            lbl_project.setOpenExternalLinks(True)
            layout.addWidget(lbl_project)

        # Close button
        self.btn_close = QPushButton("Close")
        self.btn_close.clicked.connect(self.close)

        layout.addWidget(self.btn_close)

        self.setLayout(layout)

class appconfigstore:
    def __init__(self):
        self.outputfile = "opus"
        self.outputfolderpath = ""
        self.folderdef = 0
        self.prefix = ""
        self.suffix = ""
        self.ifilexist = 0

    def save(self):
        settings = {
            'outputExt': self.outputfile,
            "outputfolderpath": self.outputfolderpath,
            "folderdef": self.folderdef,
            "prefix": self.prefix,
            "suffix": self.suffix,
            "iffileexist": self.ifilexist
        }

        with open(f"config.json", 'w') as jsonfile:
            json.dump(settings, jsonfile, indent=4)

    def load(self, ui):
        with open(f"config.json", 'r') as jsonfile:
            settings = json.load(jsonfile)
            # set Extension File
            self.outputfile = settings.get('outputExt', self.outputfile)
            if self.outputfile == "opus":
                ui.UseOpusfile.setChecked(True)
            elif self.outputfile == "ogg":
                ui.UseOGGfile.setChecked(True)
            elif self.outputfile == "oga":
                ui.UseOGAfile.setChecked(True)
            elif self.outputfile == "ts":
                ui.UseTSfile.setChecked(True)
            elif self.outputfile == "mka":
                ui.UseMKAfile.setChecked(True)
            elif self.outputfile == "caf":
                ui.UseCAFfile.setChecked(True)
            else:
                ui.UseOpusfile.setChecked(True)
            self.outputfolderpath = settings.get('outputfolderpath', self.outputfolderpath)
            ui.OutFolPath.setText(self.outputfolderpath)
            self.folderdef = settings.get('folderdef', self.folderdef)

            if self.folderdef in [1, 2, 3, 4]:
                ui.OutFolPath.setEnabled(True)
                ui.ShowFolButton.setEnabled(True)
                ui.LocaButton.setEnabled(True)
            else:
                ui.OutFolPath.setEnabled(False)
                ui.ShowFolButton.setEnabled(False)
                ui.LocaButton.setEnabled(False)

            ui.defolderIn.setCurrentIndex(self.folderdef)
            self.prefix = settings.get('prefix', self.prefix)
            ui.PrefixIn.setText(self.prefix)
            self.suffix = settings.get('suffix', self.suffix)
            ui.SuffixIn.setText(self.suffix)
            self.ifilexist = settings.get('iffileexist', self.ifilexist)

            if self.ifilexist == 0:
                ui.EREN.setChecked(True)
            elif self.ifilexist == 1:
                ui.ESKIP.setChecked(True)
            elif self.ifilexist == 2:
                ui.EOVER.setChecked(True)
            else:
                ui.EREN.setChecked(True)

def make_dual_mono(pcm, channels):
    pcm = pcm.astype(np.float32)
    if channels == 2:
        pcm = pcm[:len(pcm) - (len(pcm) % 2)]
        stereo = pcm.reshape(-1, 2)
        mono = stereo.mean(axis=1)
    else:
        mono = pcm
    mono = np.clip(mono, -32768, 32767).astype(np.int16)
    return np.column_stack((mono, mono)).flatten()

def rescale(value, in_min=0.01, in_max=1.0, out_min=0, out_max=255):
    return int((value - in_min) / (in_max - in_min) * (out_max - out_min) + out_min)


class Worker(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(int)
    processingstatus = pyqtSignal(str)

    def __init__(self, filelist, opusettings, appconfig, mainself):
        super().__init__()
        self.filelist = filelist
        self.opusettings = opusettings
        self.appconfig = appconfig
        self.mainself = mainself

        self.currentfilelen = 0
        self.currentconvertlen = 0

    def detect_max_freq_response(self, fft_result, freq, noise_threshold_db):
        """
        Detect the maximum frequency with significant energy above the noise threshold.
        Uses a more sophisticated method to find actual signal rolloff.

        Args:
            fft_result: FFT result (complex array)
            freq: Frequency bins corresponding to FFT result

        Returns:
            Maximum frequency in Hz with significant energy
        """
        # Convert to magnitude (linear scale)
        magnitude = np.abs(fft_result)

        # Convert to dB scale
        # Add small epsilon to avoid log(0)
        magnitude_db = 20 * np.log10(magnitude + 1e-10)

        # Find the peak magnitude (reference level)
        peak_db = np.max(magnitude_db)

        # Calculate relative threshold (dB below peak)
        # Looking for frequencies that are within a certain range of the peak
        relative_threshold = peak_db + noise_threshold_db

        # Find frequencies above the relative threshold
        above_threshold = magnitude_db > relative_threshold

        if np.any(above_threshold):
            # Find the highest frequency above threshold
            max_freq_idx = np.where(above_threshold)[0][-1]
            max_freq = freq[max_freq_idx]

            # Optional: Add some debug info
            if hasattr(self, 'debug_mode') and self.debug_mode:
                print(f"Peak: {peak_db:.1f} dB, Threshold: {relative_threshold:.1f} dB")

            return int(max_freq)
        else:
            return 0

    def do_work(self):
        self.processingstatus.emit("importing pyogg and libopus")
        os.environ["pyogg_win_libopus_version"] = self.opusettings.getopusstrversion()
        #if self.opusettings.getopusstrversion() == "old":
        #    os.environ["pyogg_win_libopus_version"] = "custom"
        #    os.environ["pyogg_win_libopus_custom_path"] = os.path.join(os.getcwd(), "libopus11.dll")

        importlib.reload(pyogg.opus)
        self.processingstatus.emit("creating encoder")
        opusencoder = pyogg.OpusBufferedEncoder()
        self.processingstatus.emit("configuring encoder")
        self.opusettings.setopusencoder(opusencoder)

        for file in self.filelist:
            base = os.path.basename(file)  # Get the base name of the path
            filename = os.path.splitext(base)[0]  # Split the base name into filename and extension, then take only filename
            self.processingstatus.emit(f"wait ffmpeg converting {filename} to int16 at tempfile")
            tempwavpath = self.converttowave(file)
            self.processingstatus.emit(f"reading {filename} at tempfile")
            wave_read = wave.open(tempwavpath, "rb")

            if self.appconfig.folderdef == 0:
                outputpath = os.path.dirname(file)
            elif self.appconfig.folderdef == 1:
                outputpath = str(Path.home() / "Desktop")
            elif self.appconfig.folderdef == 2:
                outputpath = str(Path.home() / "Music")
            elif self.appconfig.folderdef == 2:
                outputpath = str(Path.home() / "Documents")
            else:
                outputpath = self.appconfig.outputfolderpath

            while True:
                # Get data from the wav file
                pcm = np.frombuffer(wave_read.readframes(1024), dtype=np.int16)

                # Check if we've finished reading the wav file
                if len(pcm) == 0:
                    break

                self.currentfilelen += len(pcm)
            wave_read.rewind()

            output_filename = outputpath + "/" + self.appconfig.prefix + filename + self.appconfig.suffix + "." + self.appconfig.outputfile

            # check if file exist
            if os.path.exists(output_filename):
                if self.appconfig.ifilexist == 0:
                    counter = 0
                    cleanpath = output_filename.split(".")

                    while os.path.exists(f"{cleanpath[0]}_{counter:03d}.{cleanpath[-1]}"):
                        counter += 1

                    output_filename = f"{cleanpath[0]}_{counter:03d}.{cleanpath[-1]}"
                elif self.appconfig.ifilexist == 1:
                    os.remove(tempwavpath)
                    continue
                elif self.appconfig.ifilexist == 2:
                    pass

            if self.appconfig.outputfile in ["opus", "ogg", "oga"]:
                writer = pyogg.OggOpusWriter(output_filename, opusencoder)
            else:
                writer = None
            self.processingstatus.emit(f"converting {filename} to opus")
            avgbitrate = []
            while True:
                # Get data from the wav file
                pcm = np.frombuffer(wave_read.readframes(1024), dtype=np.int16)

                # Check if we've finished reading the wav file
                if len(pcm) == 0:
                    break

                if self.opusettings.ai:
                    mono_data = np.mean(pcm.reshape(-1, 2), axis=1)

                    fft_result = np.fft.rfft(mono_data)
                    freq = np.fft.rfftfreq(len(mono_data), 1 / self.opusettings.samplesrates)

                    # Detect max frequency response
                    max_freq = self.detect_max_freq_response(fft_result, freq, self.opusettings.absthreshold)

                    if self.opusettings.dabs:
                        bitrate = rescale(min(max_freq, 20000), 0, 20000, 2500, self.opusettings.bitrate * 1000)
                    else:
                        if self.opusettings.oldcodebook:
                            if max_freq > 18000:
                                bitrate = self.opusettings.bitrate
                            elif 14000 < max_freq <= 18000:
                                bitrate = self.opusettings.bitrate / 2
                            elif 10000 < max_freq <= 14000:
                                bitrate = self.opusettings.bitrate / 3
                            elif 6000 < max_freq <= 10000:
                                bitrate = self.opusettings.bitrate / 4
                            elif 4000 < max_freq <= 6000:
                                bitrate = self.opusettings.bitrate / 5
                            elif 2000 < max_freq <= 4000:
                                bitrate = self.opusettings.bitrate / 6
                            else:
                                bitrate = self.opusettings.bitrate / 7
                        else:
                            if max_freq > 19000:
                                bitrate = self.opusettings.bitrate
                            elif 18000 < max_freq <= 19000:
                                bitrate = self.opusettings.bitrate / 1.25
                            elif 17000 < max_freq <= 18000:
                                bitrate = self.opusettings.bitrate / 1.5
                            elif 16000 < max_freq <= 17000:
                                bitrate = self.opusettings.bitrate / 1.75
                            elif 15000 < max_freq <= 16000:
                                bitrate = self.opusettings.bitrate / 2
                            elif 14000 < max_freq <= 15000:
                                bitrate = self.opusettings.bitrate / 2.25
                            elif 13000 < max_freq <= 14000:
                                bitrate = self.opusettings.bitrate / 2.5
                            elif 12000 < max_freq <= 13000:
                                bitrate = self.opusettings.bitrate / 2.75
                            elif 11000 < max_freq <= 12000:
                                bitrate = self.opusettings.bitrate / 3.25
                            elif 10000 < max_freq <= 11000:
                                bitrate = self.opusettings.bitrate / 3.5
                            elif 9000 < max_freq <= 10000:
                                bitrate = self.opusettings.bitrate / 3.75
                            elif 8000 < max_freq <= 9000:
                                bitrate = self.opusettings.bitrate / 4
                            elif 7000 < max_freq <= 8000:
                                bitrate = self.opusettings.bitrate / 4.5
                            elif 6000 < max_freq <= 7000:
                                bitrate = self.opusettings.bitrate / 5
                            elif 5000 < max_freq <= 6000:
                                bitrate = self.opusettings.bitrate / 5.5
                            elif 4000 < max_freq <= 5000:
                                bitrate = self.opusettings.bitrate / 6
                            elif 3000 < max_freq <= 4000:
                                bitrate = self.opusettings.bitrate / 6.5
                            elif 2000 < max_freq <= 3000:
                                bitrate = self.opusettings.bitrate / 7
                            elif 1000 < max_freq <= 2000:
                                bitrate = self.opusettings.bitrate / 7.5
                            else:
                                bitrate = self.opusettings.bitrate / 8

                        bitrate = int(max(2.5, bitrate) * 1000)

                    if self.opusettings.automono and wave_read.getnchannels() == 2:
                        left_channel = pcm[::2]
                        right_channel = pcm[1::2]

                        mid = (left_channel + right_channel) / 2
                        side = (left_channel - right_channel) / 2

                        try:
                            loudnessside = 20 * math.log10(np.sqrt(np.mean(np.square(side.astype(np.float64)))) / 32768)
                        except:
                            loudnessside = 0

                        if loudnessside < self.opusettings.amthreshold:
                            # convert to mono from gainedpcm
                            output = make_dual_mono(pcm, wave_read.getnchannels())
                            bitrate = bitrate / 2
                        else:
                            output = pcm
                    else:
                        output = pcm

                    if self.opusettings.absavg:
                        avgbitrate.append(int(bitrate))
                        bitrate = int(sum(avgbitrate) / len(avgbitrate))

                    writer._encoder.set_bitrates(bitrate)
                else:
                    output = pcm

                # Encode the PCM data
                writer.write(memoryview(bytearray(output)))

                self.currentconvertlen += len(output)

                self.progress.emit((self.currentconvertlen) * 100 // self.currentfilelen)  # Emit progress value
            writer.close()
            wave_read.close()
            self.currentfilelen = 0
            self.currentconvertlen = 0
            os.remove(tempwavpath)

        self.finished.emit()

    def converttowave(self, file):
        # Create a temporary directory to store intermediate files
        temp_dir = tempfile.mkdtemp()

        base = os.path.basename(file)  # Get the base name of the path
        filename = os.path.splitext(base)[0]
        # Temporary WAV file path
        temp_wav_file = os.path.join(temp_dir, filename + ".wav")

        # Run ffmpeg to extract audio from the video file
        subprocess.run(["ffmpeg", "-i", file, "-vn", "-acodec", "pcm_s16le", "-ar", str(self.opusettings.samplesrates), "-ac", str(self.opusettings.channel), temp_wav_file], check=True)

        # Return the path to the temporary WAV file
        return temp_wav_file

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = window.Ui_MainWindow()
        self.ui.setupUi(self)
        self.opusconfig = opusconfigstorer()
        self.appconfig = appconfigstore()

        self.setWindowTitle("qOpus Encoder v1.5")

        try:
            self.appconfig.load(self.ui)
        except:
            self.appconfig.save()

        self.ui.actionAbout_3.triggered.connect(self.show_about_dialog)
        self.ui.actionExit_3.triggered.connect(self.exitbymenu)

        self.ui.PresetsIn.setPlaceholderText("Custom")

        files = os.listdir("./presets")
        file_names_without_extension = [os.path.splitext(os.path.basename(file))[0] for file in files if os.path.isfile(os.path.join("./presets", file))]
        for file_name in file_names_without_extension:
            self.ui.PresetsIn.addItem(file_name)

        self.setFixedSize(1010, 750)

        self.setAcceptDrops(True)
        self.ui.FileTable.setAcceptDrops(True)
        self.ui.FileTable.setDragEnabled(True)
        self.ui.FileTable.setDragDropMode(QAbstractItemView.DragDrop)
        self.ui.FileTable.setContextMenuPolicy(3)  # ContextMenuPolicy.CustomContextMenu
        self.ui.FileTable.customContextMenuRequested.connect(self.FileTable_show_context_menu)

        self.ui.PresetsIn.currentIndexChanged.connect(self.loadopuspreset)
        self.ui.SaveOpusPreButton.clicked.connect(self.saveopuspreset)
        self.ui.DeleteOpusPreButton.clicked.connect(self.deleteopuspreset)

        self.ui.PrefixIn.textChanged.connect(self.setpresuffix)
        self.ui.SuffixIn.textChanged.connect(self.setpresuffix)

        self.ui.EREN.clicked.connect(self.setfilexistaction)
        self.ui.ESKIP.clicked.connect(self.setfilexistaction)
        self.ui.EOVER.clicked.connect(self.setfilexistaction)

        self.ui.defolderIn.currentIndexChanged.connect(self.selectfolderdef)
        self.ui.OutFolPath.textChanged.connect(self.setoutputpath)
        self.ui.LocaButton.clicked.connect(self.selectoutputpath)
        self.ui.ShowFolButton.clicked.connect(self.showoutputpath)

        self.ui.pushButton.clicked.connect(self.selectaudiofile)
        self.ui.pushButton_2.clicked.connect(self.clear_table)

        # Opus Version select
        self.ui.UseHEV2opus.clicked.connect(self.checkversion)
        self.ui.UseHEopus.clicked.connect(self.checkversion)
        self.ui.UseNEWopus.clicked.connect(self.checkversion)
        self.ui.UseSTABLEopus.clicked.connect(self.checkversion)
        self.ui.UseOLDopus.clicked.connect(self.checkversion)
        # Opus Framesizes select
        self.ui.Use120ms.clicked.connect(self.checkframesize)
        self.ui.Use100ms.clicked.connect(self.checkframesize)
        self.ui.Use80ms.clicked.connect(self.checkframesize)
        self.ui.Use60ms.clicked.connect(self.checkframesize)
        self.ui.Use40ms.clicked.connect(self.checkframesize)
        self.ui.Use20ms.clicked.connect(self.checkframesize)
        self.ui.Use10ms.clicked.connect(self.checkframesize)
        self.ui.Use5ms.clicked.connect(self.checkframesize)
        #self.ui.Use2d5ms.clicked.connect(self.checkframesize)
        # Opus Application select
        self.ui.UseHybridApp.clicked.connect(self.checkapp)
        self.ui.UseSILKApp.clicked.connect(self.checkapp)
        self.ui.UseCELTApp.clicked.connect(self.checkapp)
        # Opus Bitrate Mode select
        self.ui.UseVBRbm.clicked.connect(self.checkbitmode)
        self.ui.UseCVBRbm.clicked.connect(self.checkbitmode)
        self.ui.UseCBRbm.clicked.connect(self.checkbitmode)
        # Opus Bandwidth select
        self.ui.UseAUTObw.clicked.connect(self.checkbandwidth)
        self.ui.UseFBbw.clicked.connect(self.checkbandwidth)
        self.ui.UseSWBbw.clicked.connect(self.checkbandwidth)
        self.ui.UseWBbw.clicked.connect(self.checkbandwidth)
        #self.ui.UseMBbw.clicked.connect(self.checkbandwidth)
        self.ui.UseNBbw.clicked.connect(self.checkbandwidth)
        # Opus Bitrate input
        self.ui.BitratesIn.valueChanged.connect(self.setbitrate)
        # Opus Channel input
        self.ui.ChannelIn.valueChanged.connect(self.setchannel)
        # Opus packet loss input
        self.ui.PACLOSIn.valueChanged.connect(self.setpackloss)
        # Opus compression input
        self.ui.CompIn.valueChanged.connect(self.setcompression)
        # Opus samples rates input
        self.ui.SampRateIn.valueChanged.connect(self.setsamprate)
        self.ui.GainIn.valueChanged.connect(self.setGain)
        self.ui.ABSTIn.valueChanged.connect(self.setABSThreshold)
        self.ui.AMTIn.valueChanged.connect(self.setAMThreshold)
        # Opus Output File Extension select
        self.ui.UseOpusfile.clicked.connect(self.checkfileextension)
        self.ui.UseOGGfile.clicked.connect(self.checkfileextension)
        self.ui.UseOGAfile.clicked.connect(self.checkfileextension)
        #self.ui.UseTSfile.clicked.connect(self.checkfileextension)
        #self.ui.UseMKAfile.clicked.connect(self.checkfileextension)
        #self.ui.UseCAFfile.clicked.connect(self.checkfileextension)
        # Opus Features select
        self.ui.EnaABS.clicked.connect(self.setFeatures)
        self.ui.EnaPred.clicked.connect(self.setFeatures)
        self.ui.EnaSPI.clicked.connect(self.setFeatures)
        self.ui.EnaDTX.clicked.connect(self.setFeatures)
        self.ui.EnaAM.clicked.connect(self.setFeatures)
        self.ui.EnaDABS.clicked.connect(self.setFeatures)
        self.ui.EnaABSAvg.clicked.connect(self.setFeatures)
        self.ui.EnaABSOCB.clicked.connect(self.setFeatures)

        self.ui.StartConvButton.clicked.connect(self.startconvert)

        self.ui.statusbar.showMessage("Ready")

    def clear_table(self):
        self.ui.FileTable.setRowCount(0)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def format_size(self, size):
        # Convert size to appropriate units (KB, MB, GB)
        for unit in ['', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.2f} {unit}"
            size /= 1024.0

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
            for url in event.mimeData().urls():
                file_path = str(url.toLocalFile())
                file_name = os.path.basename(file_path)
                file_size = os.path.getsize(file_path)
                row_position = self.ui.FileTable.rowCount()
                self.ui.FileTable.insertRow(row_position)
                self.ui.FileTable.setItem(row_position, 0, QTableWidgetItem(file_name))
                self.ui.FileTable.setItem(row_position, 1, QTableWidgetItem(file_path))
                self.ui.FileTable.setItem(row_position, 2, QTableWidgetItem(self.format_size(file_size)))
        else:
            super().dropEvent(event)

    def get_file_paths(self):
        file_paths = []
        for row in range(self.ui.FileTable.rowCount()):
            item = self.ui.FileTable.item(row, 1)  # Assuming paths are in the second column
            if item:
                file_paths.append(item.text())
        return file_paths

    def delete_row(self):
        row = self.ui.FileTable.currentRow()
        if row >= 0:
            self.ui.FileTable.removeRow(row)

    def FileTable_show_context_menu(self, point):
        menu = QMenu()
        delete_action = QAction("Delete", self)
        delete_action.triggered.connect(self.delete_row)
        menu.addAction(delete_action)
        menu.exec_(self.ui.FileTable.mapToGlobal(point))

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Delete or event.key() == Qt.Key_Backspace:
            self.delete_row()
        else:
            super().keyPressEvent(event)

    def selectaudiofile(self):
        options = QFileDialog.Options()
        file_names, _ = QFileDialog.getOpenFileNames(self, "Open Files", "", "All Files (*)", options=options)
        if file_names:
            for file_path in file_names:
                file_name = os.path.basename(file_path)
                file_size = os.path.getsize(file_path)
                row_position = self.ui.FileTable.rowCount()
                self.ui.FileTable.insertRow(row_position)
                self.ui.FileTable.setItem(row_position, 0, QTableWidgetItem(file_name))
                self.ui.FileTable.setItem(row_position, 1, QTableWidgetItem(file_path))
                self.ui.FileTable.setItem(row_position, 2, QTableWidgetItem(self.format_size(file_size)))

    def checkversion(self):
        if self.ui.UseHEV2opus.isChecked():
            self.ui.Use120ms.setEnabled(True)
            self.ui.Use100ms.setEnabled(True)
            self.ui.Use80ms.setEnabled(True)
            self.opusconfig.version = 1
        else:
            self.ui.Use120ms.setEnabled(False)
            self.ui.Use100ms.setEnabled(False)
            self.ui.Use80ms.setEnabled(False)
            if self.opusconfig.framesizes in [120, 100, 80]:
                self.opusconfig.framesizes = 60
                self.ui.Use60ms.setChecked(True)

            if self.ui.UseHEopus.isChecked():
                self.opusconfig.version = 2
            elif self.ui.UseNEWopus.isChecked():
                self.opusconfig.version = 3
            elif self.ui.UseSTABLEopus.isChecked():
                self.opusconfig.version = 4
            elif self.ui.UseOLDopus.isChecked():
                self.opusconfig.version = 5
            else:
                self.opusconfig.version = 4

    def checkframesize(self):
        if self.ui.Use120ms.isChecked():
            self.opusconfig.framesizes = 120
        elif self.ui.Use100ms.isChecked():
            self.opusconfig.framesizes = 100
        elif self.ui.Use80ms.isChecked():
            self.opusconfig.framesizes = 80
        elif self.ui.Use60ms.isChecked():
            self.opusconfig.framesizes = 60
        elif self.ui.Use40ms.isChecked():
            self.opusconfig.framesizes = 40
        elif self.ui.Use20ms.isChecked():
            self.opusconfig.framesizes = 20
        elif self.ui.Use10ms.isChecked():
            self.opusconfig.framesizes = 10
        elif self.ui.Use5ms.isChecked():
            self.opusconfig.framesizes = 5
        elif self.ui.Use2d5ms.isChecked():
            self.opusconfig.framesizes = 2.5
        else:
            self.opusconfig.framesizes = 60

    def checkapp(self):
        if self.ui.UseHybridApp.isChecked():
            self.opusconfig.app = "audio"
        elif self.ui.UseSILKApp.isChecked():
            self.opusconfig.app = "voip"
        elif self.ui.UseCELTApp.isChecked():
            self.opusconfig.app = "restricted_lowdelay"
        else:
            self.opusconfig.app = "restricted_lowdelay"

    def checkbitmode(self):
        if self.ui.UseVBRbm.isChecked():
            self.opusconfig.bitmode = "VBR"
        elif self.ui.UseCVBRbm.isChecked():
            self.opusconfig.bitmode = "CVBR"
        elif self.ui.UseCBRbm.isChecked():
            self.opusconfig.bitmode = "CBR"
        else:
            self.opusconfig.bitmode = "CVBR"

    def checkbandwidth(self):
        if self.ui.UseAUTObw.isChecked():
            self.opusconfig.bandwidth = "auto"
        elif self.ui.UseFBbw.isChecked():
            self.opusconfig.bandwidth = "fullband"
        elif self.ui.UseSWBbw.isChecked():
            self.opusconfig.bandwidth = "superwideband"
        elif self.ui.UseWBbw.isChecked():
            self.opusconfig.bandwidth = "wideband"
        elif self.ui.UseMBbw.isChecked():
            self.opusconfig.bandwidth = "mediumband"
        elif self.ui.UseNBbw.isChecked():
            self.opusconfig.bandwidth = "narrowband"
        else:
            self.opusconfig.bandwidth = "fullband"

    def setbitrate(self, value):
        self.opusconfig.bitrate = value

    def setchannel(self, value):
        self.opusconfig.channel = value

    def setpackloss(self, value):
        self.opusconfig.packloss = value

    def setcompression(self, value):
        self.opusconfig.compression = value

    def setsamprate(self, value):
        self.opusconfig.samplesrates = value

    def setGain(self, value):
        self.opusconfig.gain = value

    def setABSThreshold(self, value):
        self.opusconfig.absthreshold = value

    def setAMThreshold(self, value):
        self.opusconfig.amthreshold = value

    def setFeatures(self):
        self.opusconfig.oldcodebook = self.ui.EnaABSOCB.isChecked()
        self.opusconfig.absavg = self.ui.EnaABSAvg.isChecked()
        self.opusconfig.dabs = self.ui.EnaDABS.isChecked()
        self.opusconfig.automono = self.ui.EnaAM.isChecked()
        self.opusconfig.ai = self.ui.EnaABS.isChecked()
        self.opusconfig.prediction = self.ui.EnaPred.isChecked()
        self.opusconfig.phaseinvert = self.ui.EnaSPI.isChecked()
        self.opusconfig.DTX = self.ui.EnaDTX.isChecked()

    def checkfileextension(self):
        if self.ui.UseOpusfile.isChecked():
            self.appconfig.outputfile = "opus"
        elif self.ui.UseOGGfile.isChecked():
            self.appconfig.outputfile = "ogg"
        elif self.ui.UseOGAfile.isChecked():
            self.appconfig.outputfile = "oga"
        elif self.ui.UseTSfile.isChecked():
            self.appconfig.outputfile = "ts"
        elif self.ui.UseMKAfile.isChecked():
            self.appconfig.outputfile = "mka"
        elif self.ui.UseCAFfile.isChecked():
            self.appconfig.outputfile = "caf"
        else:
            self.appconfig.outputfile = "opus"

        self.appconfig.save()

    def startconvert(self):
        self.setEnabled(False)
        self.ui.StartConvButton.setText("Converting")
        self.ui.StartConvButton.setStyleSheet("background-color: rgb(255, 255, 0); color: rgb(0, 0, 0);")
        listfile = self.get_file_paths()
        # check if list is no file
        if not listfile:
            self.setEnabled(True)
            self.ui.statusbar.showMessage("No Files in List")
            self.ui.StartConvButton.setText("Error")
            self.ui.StartConvButton.setStyleSheet("background-color: rgb(255, 0, 0); color: rgb(0, 0, 0);")
            QApplication.processEvents()
            time.sleep(1)
            self.ui.statusbar.showMessage("Ready")
            self.ui.StartConvButton.setText("Start Convert")
            self.ui.StartConvButton.setStyleSheet("background-color: rgb(0, 255, 0); color: rgb(0, 0, 0);")

        self.worker = Worker(listfile, self.opusconfig, self.appconfig, self)
        self.worker_thread = QThread()
        self.worker.moveToThread(self.worker_thread)
        self.worker.finished.connect(self.converted)
        self.worker.progress.connect(self.updateprogress)  # Connect progress signal
        self.worker.processingstatus.connect(self.processingstatus)
        self.worker_thread.started.connect(self.worker.do_work)
        self.worker_thread.start()

    def converted(self):
        self.worker_thread.quit()
        self.worker_thread.wait()
        self.setEnabled(True)
        self.ui.statusbar.showMessage(f"converted")
        self.ui.StartConvButton.setText("Start Convert")
        self.ui.StartConvButton.setStyleSheet("background-color: rgb(0, 255, 0); color: rgb(0, 0, 0);")
        QApplication.processEvents()
        time.sleep(1)
        self.ui.statusbar.showMessage("Ready")

    def processingstatus(self, status):
        self.ui.statusbar.showMessage(status)

    def updateprogress(self, value):
        self.ui.ConvProgBar.setValue(value)

    def loadopuspreset(self):
        self.opusconfig.load(self.ui.PresetsIn.currentText(), self.ui)

    def saveopuspreset(self):
        self.opusconfig.save(self.ui.PresetsIn.currentText())
        self.ui.PresetsIn.addItem(self.ui.PresetsIn.currentText())

    def deleteopuspreset(self):
        self.opusconfig.remove(self.ui.PresetsIn.currentText())
        self.ui.PresetsIn.removeItem(self.ui.PresetsIn.currentIndex())

    def setoutputpath(self, path):
        self.appconfig.outputfolderpath = path
        self.appconfig.save()

    def selectoutputpath(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        self.ui.OutFolPath.setText(folder_path)

    def selectfolderdef(self):
        index = self.ui.defolderIn.currentIndex()

        self.appconfig.folderdef = index
        self.appconfig.save()

        if index in [1, 2, 3, 4]:
            self.ui.OutFolPath.setEnabled(True)
            self.ui.ShowFolButton.setEnabled(True)
            self.ui.LocaButton.setEnabled(True)
        else:
            self.ui.OutFolPath.setEnabled(False)
            self.ui.ShowFolButton.setEnabled(False)
            self.ui.LocaButton.setEnabled(False)

    def setpresuffix(self):
        self.appconfig.prefix = self.ui.PrefixIn.text()
        self.appconfig.suffix = self.ui.SuffixIn.text()
        self.appconfig.save()

    def showoutputpath(self):
        if self.appconfig.folderdef == 1:
            outputpath = str(Path.home() / "Desktop")
        elif self.appconfig.folderdef == 2:
            outputpath = str(Path.home() / "Music")
        elif self.appconfig.folderdef == 2:
            outputpath = str(Path.home() / "Documents")
        else:
            outputpath = self.appconfig.outputfolderpath

        if sys.platform.startswith('win'):
            # Windows
            subprocess.Popen(f'explorer "{outputpath}"', shell=True)
        elif sys.platform.startswith('darwin'):
            # MacOS
            subprocess.Popen(['open', outputpath])
        elif sys.platform.startswith('linux'):
            # Linux
            subprocess.Popen(['xdg-open', outputpath])
        else:
            print("Unsupported operating system")

    def setfilexistaction(self):
        if self.ui.EREN.isChecked():
            self.appconfig.ifilexist = 0
        elif self.ui.ESKIP.isChecked():
            self.appconfig.ifilexist = 1
        elif self.ui.EOVER.isChecked():
            self.appconfig.ifilexist = 2
        else:
            self.appconfig.ifilexist = 0
        self.appconfig.save()

    def show_about_dialog(self):
        about_dialog = AboutDialog(self)
        about_dialog.exec_()

    def exitbymenu(self):
        self.close()

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
