from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5 import uic #<---- untuk membuka GUI
import icons_rc #<----- untuk membaca icon yang ada di pyqt
import MySQLdb #untuk menyambungkan ke database
import sys #untuk mengoperasikan GUI
import datetime
from xlrd import *
from xlsxwriter import *
counter = 0

#==========================================================================
#buat screen loading

class login(QMainWindow):
	def __init__(self):
		QMainWindow.__init__(self)
		uic.loadUi('main_login.ui',self)
		#button
		self.log_in.clicked.connect(self.login_func)
		self.reg.clicked.connect(self.for_regis)

	def login_func(self):
		self.db = MySQLdb.connect(host='localhost' ,
		                          user='obaja' ,
		                          password='bajaobaja' ,
		                          database='o-lib'
		                          )
		self.cur = self.db.cursor()

		username = self.inuser.text()
		password = self.inpsw.text()

		sql = '''SELECT * FROM akun '''
		self.cur.execute(sql)
		data = self.cur.fetchall()
		for row in data:
			if username == row[1] and password == row[2]:
				self.window = MainWindow()
				self.window.show()
				self.close()
			else:
				self.error.setText('Username & Password invalid')

	def for_regis(self):
		self.regis = register()
		self.regis.show()
		self.close()
#==========================================================================
#register form

class register(QMainWindow):
	def __init__(self):
		QMainWindow.__init__(self)
		uic.loadUi('main_register.ui',self)
		self.regis.clicked.connect(self.register_func)
		self.back.clicked.connect(self.back_to_login)

	def back_to_login(self):
		self.lgin = login()
		self.lgin.show()
		self.close()

	def register_func(self):
		self.db = MySQLdb.connect(host='localhost' ,
		                          user='obaja' ,
		                          password='bajaobaja' ,
		                          database='o-lib'
		                          )
		self.cur = self.db.cursor()

		username = self.usr.text()
		password = self.psw.text()
		email = self.email.text()
		password2 = self.vpsw.text()

		if username == ('') or password == ('') or email == (''):
			self.error_regis.setText('Lengkapi data!')
		else:
			if password == password2:
				sql = '''SELECT username,email FROM akun '''
				self.cur.execute(sql)
				data = self.cur.fetchall()
				print(data)
				for row in data:
					if username == row[0] or email == row[0]:
						return self.error_regis.setText ('Sudah terdaftar!')

					else:
						self.cur.execute ( '''INSERT INTO akun(username,password,email)
						                        VALUES (%s,%s,%s)''' , (username , password , email) )
						self.db.commit()
						QMessageBox.about(self,'Registration','Berhasil melakukan registrasi')
						self.usr.setText('')
						self.psw.setText('')
						self.email.setText('')
						self.vpsw.setText('')
						break
			else :
				self.error_regis.setText ( 'Invalid!' )

#==========================================================================
#Main Apps

###########################################################################
# ===> Utama

class MainWindow(QMainWindow):
	def __init__(self):
		QMainWindow.__init__(self)
		uic.loadUi('o-lib.ui',self)
		self.UI_func()
		self.buttons()
		self.lihat_data()
		self.lihat_buku()


	def UI_func(self):
		self.mainbody.tabBar().setVisible(False)


	def buttons(self):
		self.operasi.clicked.connect(self.open_pinjam)
		self.akun.clicked.connect(self.slide_akun)
		self.slice.clicked.connect(self.slide_akun)
		self.buku.clicked.connect(self.open_buku)
		self.add.clicked.connect(self.main_pinjam)
		self.add_buku.clicked.connect(self.tambah_buku)
		self.log_2.clicked.connect(self.change_pw)
		self.log.clicked.connect(self.log_out)
		self.info.clicked.connect(self.infor)
		self.print.clicked.connect(self.export_pinjam)
		self.printbuku.clicked.connect(self.export_buku)
		self.slide.clicked.connect(self.slide_menu)

	def open_pinjam(self):
		self.mainbody.setCurrentIndex(0)

	def open_buku(self):
		self.mainbody.setCurrentIndex(1)

	def change_pw(self):
		self.psw = change()
		self.psw.show()
		self.close()

	def log_out(self):
		self.logout = login()
		self.logout.show()
		self.close()

	def infor(self):
		QMessageBox.about(self,"Information","TERIMA KASIH TELAH MENGGUNAKAN <strong>O-LIBRARY</strong>, semoga berguna untuk kedepannya")

###########################################################################
# ===> menu pinjam

	def main_pinjam(self):
		judul = self.judulbook.text()
		email = self.user.text()
		tipe = self.tipebox.currentText()
		tipe_index = self.tipebox.currentIndex()
		hari = self.haribox.currentIndex()
		minjam = datetime.date.today()
		kembali = minjam + datetime.timedelta(days=hari)

		self.db = MySQLdb.connect(host='localhost' ,
		                            user='obaja' ,
		                            password='bajaobaja' ,
		                            database='o-lib'
		                            )
		self.cur = self.db.cursor()

		self.cur.execute('''SELECT judul FROM buku WHERE judul=%s''',(judul,))
		data = self.cur.fetchone()
		if data == None:
			self.statusBar().showMessage('Buku tidak tersedia!')
		else:
			self.cur.execute('''SELECT email FROM akun WHERE email=%s''',(email,))
			dat = self.cur.fetchone()
			if dat == None:
				self.statusBar().showMessage('Email tidak terdaftar!')
			else:
				if tipe_index == 1:
					self.cur.execute ( '''INSERT INTO peminjaman(judul,email,tipe,hari,dari,sampai)
					VALUES (%s,%s,%s,%s,%s,%s)''' , (judul , email , tipe , hari , minjam , kembali) )

					self.db.commit()
					self.statusBar().showMessage("Data telah diupdate", 2000)
					self.lihat_data()

					sql = '''DELETE FROM buku WHERE judul = %s'''
					self.cur.execute(sql,([judul]))
					self.db.commit()
					self.lihat_buku()
				else:
					self.cur.execute ('''INSERT INTO peminjaman(judul,email,tipe) VALUES (%s,%s,%s)''' ,
					                   (judul,email,tipe))

					self.db.commit()
					self.statusBar().showMessage ( "Data telah diupdate" , 2000 )
					self.lihat_data()

					sql = '''DELETE FROM buku WHERE judul = %s'''
					self.cur.execute(sql,([judul]))
					self.db.commit()
					self.lihat_buku()


	def lihat_data(self):
		self.db = MySQLdb.connect(host='localhost',
		                            user='obaja',
		                            password='bajaobaja',
		                            database='o-lib'
		                            )
		self.cur = self.db.cursor()

		self.cur.execute('''SELECT judul,email,tipe,hari,dari,sampai FROM peminjaman''')
		data = self.cur.fetchall()

		self.tableWidget.setRowCount(0)
		self.tableWidget.insertRow(0)

		for row, form in enumerate(data):
			for column, item in enumerate(form):
				self.tableWidget.setItem(row, column,QTableWidgetItem(str(item)))
				column += 1

			row_position = self.tableWidget.rowCount()
			self.tableWidget.insertRow(row_position)

###########################################################################
# ===> menu buku

	def lihat_buku(self):
		self.db = MySQLdb.connect(host='localhost',
		                            user='obaja',
		                            password='bajaobaja',
		                            database='o-lib'
		                            )
		self.cur = self.db.cursor()

		self.cur.execute ('''SELECT judul,code, kategori, penulis, penerbit, harga FROM buku''')
		data = self.cur.fetchall()

		self.tableWidget_2.setRowCount(0)
		self.tableWidget_2.insertRow(0)

		for row, form in enumerate(data):
			for column, item in enumerate(form):
				self.tableWidget_2.setItem(row, column,QTableWidgetItem(str(item)))
				column += 1

			row_position = self.tableWidget_2.rowCount()
			self.tableWidget_2.insertRow(row_position)

		self.db.close()

	def tambah_buku(self):
		self.db = MySQLdb.connect(host='localhost',
		                            user='obaja',
		                            password='bajaobaja',
		                            database='o-lib'
		                            )
		self.cur = self.db.cursor()

		judul = self.judul_.text()
		code = self.code_.text()
		kategori = self.kategori_.currentText()
		penulis = self.penulis_.text()
		penerbit = self.penerbit_.text()
		harga = self.harga_.text()

		if judul == ('') or code == ('') or penulis == ('') or penerbit == ('') or harga == (''):
			self.statusBar().showMessage("Lengkapi Data!",2000)
		else:
			self.cur.execute (
				'''INSERT INTO buku(judul,code, kategori, penulis, penerbit, harga) VALUES (%s,%s,%s,%s,%s,%s)'''
				, (judul , code , kategori , penulis , penerbit , harga))

			self.db.commit ( )
			self.statusBar ( ).showMessage ( f"buku {judul} telah ditambahkan!",2000 )
			self.judul_.setText ( '' )
			self.code_.setText ( '' )
			self.kategori_.setCurrentIndex ( 0 )
			self.penulis_.setText ( '' )
			self.penerbit_.setText ( '' )
			self.harga_.setText ( '' )
			self.lihat_buku ( )

###########################################################################
# ===> print data ke excel

	def export_pinjam(self) :
		self.db = MySQLdb.connect ( host='localhost' ,
			                            user='obaja' ,
			                            password='bajaobaja' ,
			                            database='o-lib'
			                            )
		self.cur = self.db.cursor ( )
		self.cur.execute ( '''SELECT judul,email,tipe,hari,dari,sampai FROM peminjaman''' )
		data = self.cur.fetchall ( )
		wb = Workbook ( 'Data_Peminjaman.xlsx' )
		sheet1 = wb.add_worksheet ( )

		sheet1.write ( 0 , 0 , 'Judul Buku' )
		sheet1.write ( 0 , 1 , 'E-mail' )
		sheet1.write ( 0 , 2 , 'SEWA/BELI' )
		sheet1.write ( 0 , 3 , 'Tanggal Peminjaman' )
		sheet1.write ( 0 , 4 , 'Tanggal Pengembalian' )

		row_number = 1
		for row in data :
			column_number = 0
			for item in row :
				sheet1.write ( row_number , column_number , str ( item ) )
				column_number += 1
			row_number += 1

		wb.close ( )
		self.statusBar ( ).showMessage ( 'Berhasil Mencetak Salinan' )

	def export_buku(self):
		self.db = MySQLdb.connect ( host='localhost' ,
			                            user='obaja' ,
			                            password='bajaobaja' ,
			                            database='o-lib'
			                            )
		self.cur = self.db.cursor ( )
		self.cur.execute ( '''SELECT judul,code, kategori, penulis, penerbit, harga FROM buku''' )
		data = self.cur.fetchall ( )

		wb = Workbook('Data Buku.xlsx')
		sheet1 = wb.add_worksheet()

		sheet1.write ( 0 , 0 , 'Judul Buku' )
		sheet1.write ( 0 , 1 , 'Code buku' )
		sheet1.write ( 0 , 2 , 'Kategori' )
		sheet1.write ( 0 , 3 , 'Penulis' )
		sheet1.write ( 0 , 4 , 'Penerbit' )
		sheet1.write(0,5, 'Harga')

		row_number = 1
		for row in data :
			column_number = 0
			for item in row :
				sheet1.write ( row_number , column_number , str ( item ) )
				column_number += 1
			row_number += 1

		wb.close ( )
		self.statusBar ( ).showMessage ( 'Berhasil Mencetak Salinan' )

###########################################################################
# ===> slide menu

	def slide_menu(self):
		width = self.framekiri.width()
		if width == 0:
			new_width = 200
		else:
			new_width = 0

		self.animation = QPropertyAnimation(self.framekiri, b"maximumWidth")
		self.animation.setDuration(250)
		self.animation.setStartValue(width)
		self.animation.setEndValue(new_width)
		self.animation.setEasingCurve(QEasingCurve.InOutQuart)
		self.animation.start()

	def slide_akun(self):
		width = self.akun_menu.width()

		if width == 0:
			new_width = 400
		else:
			new_width = 0


		self.animation=QPropertyAnimation(self.akun_menu, b"maximumWidth")
		self.animation.setDuration(250)
		self.animation.setStartValue(width)
		self.animation.setEndValue(new_width)
		self.animation.setEasingCurve(QEasingCurve.InOutQuart)
		self.animation.start()


#==========================================================================
#edit password

class change (QMainWindow):
	def __init__(self):
		QMainWindow.__init__(self)
		uic.loadUi('password.ui',self)
		self.go.clicked.connect(self.change_password)
		self.change.clicked.connect(self.new_pass)
		self.changemenu.hide()



	def change_password(self):
		self.db = MySQLdb.connect ( host='localhost' ,
		                            user='obaja' ,
		                            password='bajaobaja' ,
		                            database='o-lib'
		                            )
		self.cur = self.db.cursor ()

		username = self.usr_3.text()
		password = self.psw_3.text()
		email = self.email_3.text()
		vpassword = self.vpsw_3.text()
		sql = '''SELECT * FROM akun '''
		self.cur.execute(sql)
		data = self.cur.fetchall()
		for row in data:
			if username == row[1] and password == row[2] and email == row[3] :
				if password == vpassword:
					self.changemenu.show()
				else:
					self.error_regis_3.setText ( 'invalid' )
			else:
				self.error_regis_3.setText('Username & Password invalid')

	def new_pass(self):
		password = self.ps.text()
		ubah = self.con.text()
		old = self.ps_2.text()

		if password == ubah:
			self.db = MySQLdb.connect (host='localhost' ,
			                            user='obaja' ,
			                            password='bajaobaja' ,
			                            database='o-lib'
			                            )
			self.cur = self.db.cursor ()
			sql = '''UPDATE akun SET password=%s WHERE password = %s'''
			self.cur.execute(sql,(password,old,))
			data = self.db.commit()
			if data == None:
				QMessageBox.about( self , "Succes" , "Password sudah diganti" )
				self.back = login()
				self.back.show()
				self.close()
			else:
				QMessageBox.warning ( self , "Error" , "Gagal Mengganti" )
		else:
			QMessageBox.warning ( self , "Error" , "Invalid!" )




#==========================================================================
#screen loading

class loading(QMainWindow):
	def __init__(self):
		QMainWindow.__init__(self)
		uic.loadUi('progress.ui',self)


		self.setWindowFlag(Qt.FramelessWindowHint)
		self.setAttribute(Qt.WA_TranslucentBackground)

		self.timer = QTimer()
		self.timer.timeout.connect(self.pro_screen)
		self.timer.start(35)
		self.show()

	def pro_screen(self):

		global counter
		self.loading.setValue(counter)
		if counter > 100:
			self.timer.stop()
			self.main = login()
			self.main.show()
			self.close()

		counter += 1


#################################
########### MACHINE #############
#################################
if __name__ == '__main__':
	app = QApplication(sys.argv)
	window = loading()
	window.show()
	app.exec_()



