# file: nopeslite-xls.py

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import sqlite3 as lite
import xlrd
import xlwt
import os
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import folio
from reportlab.lib.units import mm


FILE_SQL = "./database.sqlite3"


class NopesliteXLS:
	def __init__(self, parent):
		self.parent = parent
		self.parent.title(":: Nopeslite - XLS ::")
		self.parent.protocol("WM_DELETE_WINDOW", self.klik_btn_pass)
		self.parent.resizable(False, False)
		
		self.atur_database()
		self.atur_komponen()
			
	def atur_database(self):
		statusDB = os.path.exists(FILE_SQL)
		
		if statusDB:
			os.remove(FILE_SQL)
		
		# set-koneksi-kursor
		self.koneksi = lite.connect(FILE_SQL)
		self.kursor = self.koneksi.cursor()
		
		self.objAturDatabase = AturDatabase(self.koneksi, self.kursor)
		self.objAturDatabase.buat_database()
		
	def atur_komponen(self):
		mainframe = tk.Frame(self.parent, bd=5)
		mainframe.pack(fill='both', expand=1)
		
		objAturKomponen = AturKomponen(self.parent, mainframe, self.koneksi, self.kursor,
						self.objAturDatabase)
		
	def klik_btn_pass(self, event=None):
		pass
		
		
class AturDatabase:
	def __init__(self, koneksi, kursor):
		self.koneksi = koneksi
		self.kursor = kursor
		
	def buat_database(self):
		sql_dataray ="""
			CREATE TABLE data_rayon(
				kdray text primary key,
				namray text)
			"""
		
		sql_datakec ="""
			CREATE TABLE data_kecamatan(
				kdkec text primary key,
				namkec text,
				kdray text)
			"""

		sql_datasek ="""
			CREATE TABLE data_sekolah(
				kdsek text primary key,
				namsek text,
				kdkec text,
				kdray text)
			"""

		sql_datasis ="""
			CREATE TABLE data_siswa(
				norut text,
				nopes text,
				nisn text,
				namsis text,
				kdsek text,
				kdkec text,
				kdray text)
			"""
			
		self.kursor.execute(sql_dataray)
		self.kursor.execute(sql_datakec)
		self.kursor.execute(sql_datasek)
		self.kursor.execute(sql_datasis)
		
		self.isi_data_rayon_kec()
		
	def isi_data_rayon_kec(self):
		# input data rayon
		DATA_RAYON = [
			('27', 'Kabupaten Malang'),
			('41', 'Kabupaten Malang')]
			
		for infoRay in DATA_RAYON:
			self.kursor.execute("INSERT INTO data_rayon VALUES (?,?)",infoRay)
			
		# input data kecamatan
		for dataRay in DATA_RAYON:
			if dataRay[0]=="27":
				DATA_KEC = [
					('01', 'PUJON'), ('02', 'KASEMBON'), ('03', 'NGANTANG'), ('04', 'SINGOSARI'),
					('05', 'LAWANG'), ('06', 'KARANGPLOSO'), ('07', 'DAU'), ('08', 'TUMPANG'),
					('09', 'PAKIS'), ('10', 'JABUNG'), ('11', 'PAKISAJI'), ('12', 'PONCOKUSUMO'),
					('13', 'BULULAWANG'), ('14', 'TAJINAN'), ('15', 'GONDANGLEGI'), ('16', 'WAJAK'),
					('17', 'TUREN'), ('18', 'DAMPIT'), ('19', 'AMPELGADING'), ('20', 'SUMBERMANJING'),
					('21', 'KEPANJEN'), ('22', 'SUMBERPUCUNG'), ('23', 'NGAJUM'), ('24', 'WAGIR'),
					('25', 'PAGAK'), ('26', 'KALIPARE')]
			elif dataRay[0]=="41":
				DATA_KEC = [
					('27', 'DONOMULYO'), ('28', 'BANTUR'), ('29', 'TIRTOYUDO'), ('30', 'GEDANGAN'),
					('31', 'WONOSARI'), ('32', 'KROMENGAN'), ('33', 'PAGELARAN')]
			
			for kdkec, namkec in DATA_KEC:
				infoKec = (kdkec, namkec, dataRay[0])
				self.kursor.execute("INSERT INTO data_kecamatan VALUES (?,?,?)",infoKec)
				
		self.koneksi.commit()
		
	def hapus_database(self):
		self.kursor.execute("DROP TABLE data_rayon")
		self.kursor.execute("DROP TABLE data_kecamatan")
		self.kursor.execute("DROP TABLE data_sekolah")
		self.kursor.execute("DROP TABLE data_siswa")
		
	def input_data_sekolah(self, datasek):
		print("input_data_sekolah")
		for data in datasek:
			kode = data[0]
			nama = data[1]
			keca = data[2]
			rayo = data[3]
			
			infoSek = (kode, nama, keca, rayo)
			#print(infoSek)
			self.kursor.execute("INSERT INTO data_sekolah VALUES (?,?,?,?)",infoSek)
			
		self.koneksi.commit()
			
	def input_data_siswa(self, datasis):
		print("input_data_siswa")
		norut = 0
		for data in datasis:
			norut += 1
			nopes = data[0]
			nisn = data[1]
			nama = data[2]
			seko = data[3]
			keca = data[4]
			rayo = data[5]
			
			infoSis = (norut, nopes, nisn, nama, seko, keca, rayo)
			#print(infoSis)
			self.kursor.execute("INSERT INTO data_siswa VALUES (?,?,?,?,?,?,?)",infoSis)
			
		self.update_nopes()
		self.koneksi.commit()
		
	def update_nopes(self):
		kodesek = self.ambil_datasek()
		#print(kodesek)
		
		for kode in kodesek:
			datasiswa = self.ambil_datasis(kode)
			#print(datasiswa)
			
			i = 0
			kode = 8
			
			for data in datasiswa:
				i += 1
				norut = data[0]
				kdsek = data[1]
				kdray = data[2]
				
				if (i%8)==0:
					kode = 9
					
				nopes = self.buat_nopes(i, kdray, kdsek, kode)
				#print(nopes)
				
				sql = "UPDATE data_siswa SET nopes=? WHERE norut=?"
				data = (nopes, norut)
				self.kursor.execute(sql, data)
				
				kode -= 1
				
	def ambil_datasek(self):
		sql = "SELECT kdsek, namsek FROM data_sekolah ORDER BY kdsek ASC"
		datasek = self.eksekusi(sql)
		
		kodesek = []
		for data in datasek:
			kodesek.append(data[0])
		
		return kodesek
		
	def ambil_datasis(self, kode):
		sql = "SELECT norut, kdsek, kdray FROM data_siswa WHERE kdsek='{}'".format(kode)
		datasis = self.eksekusi(sql)
		
		datsis = []
		for data in datasis:
			datsis.append(data)
			
		return datsis
		
	def ambil_datasis_all(self, kodesek=None):
		if kodesek==None:
			sql = """
					SELECT a.norut, a.nopes, a.nisn, a.namsis, b.namsek
					FROM data_siswa a, data_sekolah b
					WHERE a.kdsek=b.kdsek
				"""
		else:
			sql = """
					SELECT a.norut, a.nopes, a.nisn, a.namsis, b.namsek
					FROM data_siswa a, data_sekolah b
					WHERE a.kdsek=b.kdsek AND a.kdsek='{}'
				""".format(kodesek)
			
		datasis = self.eksekusi(sql)
		
		datsis = []
		for data in datasis:
			norut = data[0]
			nopes = data[1]
			nisn = data[2]
			namsis = data[3]
			namsek = data[4]
			
			infosis = (norut, nopes, nisn, namsis, namsek)
			
			datsis.append(infosis)
			
		return datsis
		
	def buat_nopes(self, i, kdray, kdsek, kode):
		if (i < 10):
			nopes = '{}-0{}-000{}-{}'.format(kdray, kdsek, i, kode)
		elif (i < 100):
			nopes = '{}-0{}-00{}-{}'.format(kdray, kdsek, i, kode)
		elif (i < 1000):
			nopes = '{}-0{}-0{}-{}'.format(kdray, kdsek, i, kode)
		
		return nopes
		
	def baca_file(self, namafile):
		# baca-file-excel
		wb = xlrd.open_workbook(namafile)
		sh = wb.sheet_by_index(0)
		
		namfile = namafile.split("/")[-1]
		kdkec = namfile[:2]

		awal = 1
		akhir = sh.nrows
		
		datasis = []
		datasek = []
		temp = ""
		
		for i in range(awal, akhir):
			kode = sh.cell(i, 1).value
			
			kdray = kode[3:5]
			kdsek = kode[7:10]
			
			nopes = "000"
			nisn = sh.cell(i, 2).value
			nama = sh.cell(i, 3).value
			namsek = sh.cell(i, 4).value
			
			infoSis = (nopes, nisn, nama.upper(), kdsek, kdkec, kdray)
			infoSek = (kdsek, namsek.upper(), kdkec, kdray)
			
			# list data-siswa
			datasis.append(infoSis)
			
			if temp != kdsek:
				temp = kdsek
				datasek.append(infoSek)
				
		return (datasek, datasis)
		
	def ambil_namakec(self, kodekec):
		sql = "SELECT namkec FROM data_kecamatan WHERE kdkec='{}'".format(kodekec)
		namakec = self.eksekusi(sql)
		
		return namakec[0][0]
		
	def ambil_namasek(self, kodesek):
		sql = "SELECT namsek FROM data_sekolah WHERE kdsek='{}'".format(kodesek)
		namasek = self.eksekusi(sql)
		
		return namasek[0][0]		
		
	def eksekusi(self, sql):
		self.kursor.execute(sql)
		data = self.kursor.fetchall()

		return data

		
class AturKomponen:
	def __init__(self, parent, frame, koneksi, kursor, db):
		self.parent = parent
		self.frame = frame
		self.koneksi = koneksi
		self.kursor = kursor
		self.db = db
		
		self.set_widget()
		self.form_load()
		
	def set_widget(self):
		tk.Label(self.frame, text="Masukkan file xls data siswa:").pack(side='top')
		
		### box1
		box1 = tk.Frame(self.frame)
		box1.pack(side='top', pady=5)
		
		self.ent_filedata = ttk.Entry(box1, width=60)
		self.ent_filedata.pack(side='left')
		
		self.btn_input = ttk.Button(box1, text='Input', command=self.klik_btn_input,
									width=5)
		self.btn_input.pack(side='left')
		
		### box2
		box2 = tk.Frame(self.frame)
		box2.pack(side='top')
		
		self.btn_rekap = ttk.Button(box2, text='Rekap XLS', command=self.klik_btn_rekap)
		self.btn_rekap.pack(side='left')
		
		self.btn_kartu = ttk.Button(box2, text='Cetak Kartu', command=self.klik_btn_cetak)
		self.btn_kartu.pack(side='left', padx=5)
		
		self.btn_keluar = ttk.Button(box2, text='Keluar', command=self.klik_btn_keluar)
		self.btn_keluar.pack(side='left')
		
	def form_load(self):
		self.btn_input.focus_set()
		
	def klik_btn_input(self):
		self.namafile = fd.askopenfilename(initialdir = "./",
			title = "Pilih File Data Siswa",
			filetypes=[('File Excel', '*.xls')],
			parent=self.parent)

		if len(self.namafile)==0:
			pass
		else:
			# hapus isi entry-file, kemudian isi sesuai file-data
			self.ent_filedata.delete(0, 'end')
			self.ent_filedata.insert('end', self.namafile)
			
			# bersihkan database
			self.db.hapus_database()
			self.db.buat_database()
			
			# baca file & simpan data ke database
			data = self.db.baca_file(self.namafile)
			
			#print(data[1])

			self.db.input_data_sekolah(data[0])
			self.db.input_data_siswa(data[1])
			
		print("+++ Validasi SUKSES! +++")
		
	def klik_btn_rekap(self):
		objEksporExcel = EksporExcel(self.db, self.namafile)
		objEksporExcel.buat_rekap()
		
	def klik_btn_cetak(self):
		objKartuUjian = KartuUjian(self.db, self.namafile)
		objKartuUjian.buat_kartu()
		
	def klik_btn_keluar(self):
		self.parent.destroy()
		self.kursor.close()
		self.koneksi.close()
		
		
class EksporExcel:
	def __init__(self, database, namafile):
		self.db = database
		self.namafile = namafile
		
	def buat_rekap(self):
		# set nama-file
		self.namfile = self.namafile.split("/")[-1]
		namasimpan = "{}_DATABASE.xls".format(self.namfile.split(".")[0])
		#print(namasimpan)
		
		bookData = xlwt.Workbook()

		self.style_excel()

		self.laporan_utama(bookData)
		
		kdsek = self.db.ambil_datasek()
		
		for kode in kdsek:
			self.laporan_sekolah(bookData, kode)

		bookData.save(namasimpan)
		print("laporan-finish...")
		
	def laporan_utama(self, bookData):
		# buat-sheet utama
		sheet = bookData.add_sheet("UTAMA")

		# atur posisi sheet
		sheet.set_portrait(True)
		
		# atur header-footer --> kosong
		sheet.set_header_str("".encode())
		sheet.set_footer_str("".encode())
		
		# atur margin
		sheet.set_left_margin(0.2)
		sheet.set_right_margin(0.2)
		sheet.set_top_margin(0.4)
		sheet.set_bottom_margin(0.4)
		
		# atur cetak tengah horizontal
		sheet.set_print_centered_horz(True)
		
		baris = 5
		no_urut = 0
		
		# ambil nama kecamatan
		kdkec = self.namfile[:2]
		nama_kec = self.db.ambil_namakec(kdkec)
		print(nama_kec)

		# kop Laporan
		sheet.write_merge(0, 0, 0, 4, 'CROSS CHECK DATA', self.stKop1)

		sheet.write(3, 0, 'KEC: {}'.format(nama_kec), self.stKop2)
			
		# buat judul kolom
		sheet.write(4, 0, 'NO', self.stJudulKuning)
		sheet.write(4, 1, 'NO PESERTA', self.stJudulKuning)
		sheet.write(4, 2, 'NISN', self.stJudulKuning)
		sheet.write(4, 3, 'NAMA PESERTA', self.stJudulKuning)
		sheet.write(4, 4, 'ASAL SEKOLAH', self.stJudulKuning)
		
		# buat lebar kolom
		sheet.col(0).width = 1440 # Nomor_Siswa
		sheet.col(1).width = 3960 # Nomor Peserta
		sheet.col(2).width = 3960 # Nomor Peserta
		sheet.col(3).width = 9000 # Nama Peserta
		sheet.col(4).width = 7300 # Asal Sekolah
		
		# ambil data_siswa
		datasiswa = self.db.ambil_datasis_all()
		norut = 0
		
		for data in datasiswa:
			norut += 1
			
			sheet.write(baris, 0, norut, self.stTengah)
			sheet.write(baris, 1, data[1], self.stTengah)
			sheet.write(baris, 2, data[2], self.stTengah)
			sheet.write(baris, 3, data[3], self.stBatas)
			sheet.write(baris, 4, data[4], self.stBatas)

			baris += 1

	def laporan_sekolah(self, bookData, kode):
		# buat-sheet utama
		sheet = bookData.add_sheet("{}".format(kode))

		# atur posisi sheet
		sheet.set_portrait(True)
		
		# atur header-footer --> kosong
		sheet.set_header_str("".encode())
		sheet.set_footer_str("".encode())
		
		# atur margin
		sheet.set_left_margin(0.2)
		sheet.set_right_margin(0.2)
		sheet.set_top_margin(0.4)
		sheet.set_bottom_margin(0.4)
		
		# atur cetak tengah horizontal
		sheet.set_print_centered_horz(True)
		
		baris = 5
		no_urut = 0
		
		# ambil nama kecamatan
		kdkec = self.namfile[:2]
		nama_kec = self.db.ambil_namakec(kdkec)
		print(kode)

		# kop Laporan
		sheet.write_merge(0, 0, 0, 4, 'DINAS PENDIDIKAN KABUPATEN MALANG', self.stKop1)
		sheet.write_merge(1, 1, 0, 4, 'DAFTAR PESERTA TRY OUT UJIAN SEKOLAH SD', self.stKop1)
		sheet.write_merge(2, 2, 0, 4, 'TAHUN AJARAN 2019/2020', self.stKop1)

		sheet.write(3, 0, 'KEC: {}'.format(nama_kec), self.stKop2)
			
		# buat judul kolom
		sheet.write(4, 0, 'NO', self.stJudulKuning)
		sheet.write(4, 1, 'NO PESERTA', self.stJudulKuning)
		sheet.write(4, 2, 'NISN', self.stJudulKuning)
		sheet.write(4, 3, 'NAMA PESERTA', self.stJudulKuning)
		sheet.write(4, 4, 'ASAL SEKOLAH', self.stJudulKuning)
		
		# buat lebar kolom
		sheet.col(0).width = 1440 # Nomor_Siswa
		sheet.col(1).width = 3960 # Nomor Peserta
		sheet.col(2).width = 3960 # Nomor Peserta
		sheet.col(3).width = 9000 # Nama Peserta
		sheet.col(4).width = 7300 # Asal Sekolah
		
		# ambil data_siswa
		datasiswa = self.db.ambil_datasis_all(kode)
		norut = 0
		
		for data in datasiswa:
			norut += 1
			
			sheet.write(baris, 0, norut, self.stTengah)
			sheet.write(baris, 1, data[1], self.stTengah)
			sheet.write(baris, 2, data[2], self.stTengah)
			sheet.write(baris, 3, data[3], self.stBatas)
			sheet.write(baris, 4, data[4], self.stBatas)

			baris += 1

	def style_excel(self):
		# style-kop
		self.stKop1 = xlwt.easyxf('alignment: horizontal center, vertical center;'
							 'font: bold true, name Calibri, height 280')
		self.stKop2 = xlwt.easyxf('alignment: horizontal left, vertical center;'
							 'font: bold true, name Calibri, height 240')
		
		# style-judul-kuning
		self.stJudulKuning = xlwt.easyxf('font: name Calibri, height 220, bold true;'
					'borders: left thin, right thin, top thin, bottom thin;'
					'pattern: pattern solid, fore_colour yellow;'
					'alignment: horizontal center, vertical center')
					
		# style-tengah
		self.stTengah = xlwt.easyxf('font: name Calibri, height 220;'
								'borders: left thin, right thin, top thin, bottom thin;'
								'alignment: horizontal center')		

		# style-batas
		self.stBatas = xlwt.easyxf('font: name Calibri, height 220;'
							'borders: left thin, right thin, top thin, bottom thin')
							
							
class KartuUjian:
	def __init__(self, database, namafile):
		self.db = database
		self.namafile = namafile
		
	def buat_kartu(self):
		# set nama-file
		self.namfile = self.namafile.split("/")[-1]
		namasimpan = "{}_KARTU.pdf".format(self.namfile.split(".")[0])

		c = Canvas(namasimpan, folio)
		self.konversi_pdf(c)
		c.save()
				
		print("+++ Konversi Data Finish +++")
		
	def konversi_pdf(self, c):
		kdsek = self.db.ambil_datasek()
		
		for kode in kdsek:
			datasiswa = self.db.ambil_datasis_all(kode)
			
			judul1 = "TRY OUT UJIAN SEKOLAH SD"
			judul2 = "TAHUN PELAJARAN 2019/2020"
			judul3 = "Kabupaten Malang, 3 Februari 2020"
			
			namasek = self.db.ambil_namasek(kode)
			namkepsek = ""
			nipkepsek = ""
			
			self.set_kartu(c, 330, datasiswa, judul1, judul2,
						namasek, judul3, namkepsek, nipkepsek)
			
			if (self.baris_akhir != 4):
				c.showPage()
				
			print(kode)
			
	def set_kartu(self, c, setY, data, judul1, judul2, namasek, tglujian, namakep, nipkep):
		jumDat = 0
		i = 0
		
		for dat in data:
			y = (setY*mm) - (i*80*mm)

			point = [6, 114]

			for idx in point:
				x = idx*mm
				
				c.rect(x, y-80*mm, 95*mm, 70*mm)
				c.line(x, y-30*mm, x+95*mm, y-30*mm)

				c.drawImage("./logo-diknas.jpg", x+5*mm, y-27.5*mm, 15*mm, 15*mm)

				c.setFont("Times-Bold", 14)
				c.drawString(x+35*mm, y-17*mm, "KARTU PESERTA")

				c.setFont("Times-Bold", 10)
				c.drawString(x+25*mm, y-22*mm, judul1)
				c.drawString(x+30*mm, y-27*mm, judul2)

				kd_pes = dat[1]
				nm_pes = dat[3]
				tmp_tgl = "-"
				sklh = namasek
				tgl_ujian = tglujian
				nm_kepsek = namakep
				nip = nipkep

				#----------------------------------
				c.setFont("Times-Bold", 9)
				c.drawString(x+28*mm, y-35*mm, ": %s" %kd_pes)

				c.setFont("Times-Roman", 9)
				c.drawString(x+3*mm, y-35*mm, "No. Peserta")
				c.drawString(x+3*mm, y-39*mm, "Nama Peserta")
				c.drawString(x+3*mm, y-43*mm, "Tmp & Tgl Lahir")
				c.drawString(x+3*mm, y-47*mm, "Sekolah Asal")

				c.drawString(x+28*mm, y-39*mm, ": %s" %nm_pes[:25])
				c.drawString(x+28*mm, y-43*mm, ": %s" %tmp_tgl)
				c.drawString(x+28*mm, y-47*mm, ": %s" %sklh)

				c.drawString(x+33*mm, y-54*mm, tgl_ujian)
				c.drawString(x+33*mm, y-58*mm, "Kepala Sekolah Penyelenggara")
				c.drawString(x+33*mm, y-71*mm, nm_kepsek)
				c.drawString(x+33*mm, y-75*mm, nip)
				
				#-------------------------------
				#logo-foto
				c.drawImage('./foto-siswa.jpg',x+8*mm, y-73*mm, 16*mm, 19*mm)

			jumDat += 1
			i += 1
			self.baris_akhir = i

			if (i==4):
				i = 0
			if (jumDat%4==0):
				c.showPage()
	

			
				
if __name__ == '__main__':
	root = tk.Tk()
	
	app = NopesliteXLS(root)
	
	root.mainloop()
		

