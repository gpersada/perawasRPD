import streamlit as st
import pandas as pd
import io
import numpy as np
from csv import writer
import datetime

buffer = io.BytesIO()
buffer2 = io.BytesIO()

st.set_page_config(page_title='[DJPb Babel] Perawas RPD')
st.title('Inovasi Perawas RPD 📋')
st.subheader('Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung')
st.write("Aplikasi Perencanaan dan Pengawasan RPD (Perawas RPD) berfungsi membantu proses penyusunan RPD Satker Anda.")
st.write("Follow [Instagram](https://www.instagram.com/djpbbabel/) dan [Youtube](https://www.youtube.com/channel/UCM70tgByAEPpIPqME2mKKyQ).")
st.write("Panduan aplikasi [klik di sini](https://drive.google.com/drive/folders/1QiATR0jJ2P4jRWGKG54gXUeqQXtM_qcA?usp=sharing).")
st.markdown('---')

uploaded_file = st.file_uploader('Upload File RPD di sini. Nama file unduhan "Form Cetak POK".', type='xlsx')
if uploaded_file:
	st.markdown('---')
	raw = pd.read_excel(uploaded_file, index_col=None, header=6, skipfooter=5, engine='openpyxl')
	info = pd.read_excel(uploaded_file, index_col=None, nrows = 1, dtype=str, engine='openpyxl')
	info.rename(columns = {'Unnamed: 4':'kdsatker', 'Unnamed: 6':'nmsatker'}, inplace = True)
	info['kdsatker'] = info['kdsatker'].str[-6:]
	nama_satker = info['nmsatker'].iloc[0]  +' (' + info['kdsatker'].iloc[0] + ')'
	kode_satker = info['kdsatker'].iloc[0]
	
	# ct stores current time
	ct = datetime.datetime.now()

	# Log RPD Terakhir 
	List=[ct,kode_satker,nama_satker]

	# Open our existing CSV file in append mode
	# Create a file object for this file
	with open('logrpdterakhir.csv', 'a') as f_object:

	    # Pass this file object to csv.writer()
	    # and get a writer object
	    writer_object = writer(f_object)

	    # Pass the list as an argument into
	    # the writerow()
	    writer_object.writerow(List)

	    #Close the file object
	    f_object.close()	

	# Master Data

	df = raw.iloc[:,[1, 2, 17, 18, 19, 20, 21, 22, 24, 26, 27, 28, 29, 31]]

	df2=df.dropna()
	df2.columns = ['Kode', 'Uraian', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

	df2=df2.loc[df2['Uraian'].str.len() > 1]
	df2[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]=df2[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]*1000

	df3=df2[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]

	# Buat ID

	idmaster = df2[['Kode', 'Uraian']]

	# Pemisahan Kode antar level RKA KL

	idprog = idmaster.loc[idmaster['Kode'].str.len() == 9]
	idgiat = idmaster.loc[idmaster['Kode'].str.len() == 4]
	idkro = idmaster.loc[idmaster['Kode'].str.len() == 8]
	idro = idmaster.loc[idmaster['Kode'].str.len() == 7]
	idkomp = idmaster.loc[idmaster['Kode'].str.len() == 3]
	idakun = idmaster.loc[idmaster['Kode'].str.len() == 6]

	idprog.columns = ['kdprog', 'Uraian']
	idgiat.columns = ['kdgiat', 'Uraian']
	idkro.columns = ['kdkro', 'Uraian']
	idro.columns = ['kdro', 'Uraian']
	idkomp.columns = ['kdkomp', 'Uraian']

	### Mengambil jenis belanja saja

	idjenbel = idakun['Kode'].str[:2]

	### Penyesuaian Kode Program, KRO, dan RO

	kdprog = idprog['kdprog'].str[-2:]

	kdkro = idkro['kdkro'].str[-3:]

	kdro = idro['kdro'].str[-3:]

	kdgiat = idgiat['kdgiat']

	### Proses Bentuk Realisasi Komponen & Jenis Belanja

	kompbel1 = pd.concat([idkomp,idjenbel], axis=1)

	kompbel1.loc[:,'kdkomp'] = kompbel1.loc[:,'kdkomp'].ffill()
	kompbel1.loc[:,'Uraian'] = kompbel1.loc[:,'Uraian'].ffill()
	kompbel1=kompbel1.dropna()
	kompbel1['ID']=kompbel1['kdkomp']+'.'+kompbel1['Kode']
	kompbel2 = kompbel1[['ID', 'Uraian', 'Kode']]
	kompbel2.rename(columns = {'Kode':'Belanja'}, inplace = True)

	kompbel3 = pd.concat([kompbel2,kdro], axis=1)
	kompbel3.loc[:,'kdro'] = kompbel3.loc[:,'kdro'].ffill()
	kompbel3=kompbel3.dropna()
	kompbel3['ID']=kompbel3['kdro']+'.'+kompbel3['ID']
	kompbel4 = kompbel3[['ID', 'Uraian', 'Belanja']]

	kompbel5 = pd.concat([kompbel4,kdkro], axis=1)
	kompbel5.loc[:,'kdkro'] = kompbel5.loc[:,'kdkro'].ffill()
	kompbel5=kompbel5.dropna()
	kompbel5['ID']=kompbel5['kdkro']+'.'+kompbel5['ID']
	kompbel6 = kompbel5[['ID', 'Uraian', 'Belanja']]

	kompbel7 = pd.concat([kompbel6,kdgiat], axis=1)
	kompbel7.loc[:,'kdgiat'] = kompbel7.loc[:,'kdgiat'].ffill()
	kompbel7=kompbel7.dropna()
	kompbel7['ID']=kompbel7['kdgiat']+'.'+kompbel7['ID']
	kompbel8 = kompbel7[['ID', 'Uraian', 'Belanja']]

	kompbel9 = pd.concat([kompbel8,kdprog], axis=1)
	kompbel9.loc[:,'kdprog'] = kompbel9.loc[:,'kdprog'].ffill()
	kompbel9=kompbel9.dropna()
	kompbel9['ID']=kompbel9['kdprog']+'.'+kompbel9['ID']
	kompbel10 = kompbel9[['ID', 'Uraian', 'Belanja']]

	rpdkomp = pd.concat([kompbel10,df3], axis=1)
	rpdkomp=rpdkomp.dropna()

	rpdkomp['Belanja'] = rpdkomp['Belanja'].replace(to_replace = ['51','52','53','57'], value = ['Pegawai','Barang','Modal','Bantuan Sosial'])

	rpdkompsum = rpdkomp.groupby(['ID','Uraian','Belanja']).sum()[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]
	rpdkompsum.reset_index(inplace=True)

	rpdkompsum['Pagu']= rpdkompsum[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']].sum(axis=1)
	rpdkompsum['Kode'] = rpdkompsum['ID'].str[16:19]
	rpdkompsum['Keterangan']=np.nan
	rpdkompsum['Sisa RPD']=np.nan
	rpdkompsum = rpdkompsum[['ID','Kode','Uraian', 'Belanja', 'Keterangan', 'Pagu', 'Sisa RPD','Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]


	### Proses Bentuk Uraian RO & KRO

	urro = pd.concat([kdprog,kdgiat,idro], axis=1)
	urro.loc[:,'kdprog'] = urro.loc[:,'kdprog'].ffill()
	urro.loc[:,'kdgiat'] = urro.loc[:,'kdgiat'].ffill()
	urro=urro.dropna()
	urro['ID']=urro['kdprog']+'.'+urro['kdgiat']+'.'+urro['kdro']+'.000.00'
	urro['Kode']=urro['kdgiat']+'.'+urro['kdro']
	urro = urro[['ID', 'Kode','Uraian']]
	urro.reset_index(drop=True, inplace=True)

	urkro = pd.concat([kdprog,idkro], axis=1)
	urkro.loc[:,'kdprog'] = urkro.loc[:,'kdprog'].ffill()
	urkro=urkro.dropna()
	urkro['ID']=urkro['kdprog']+'.'+urkro['kdkro']+'.000.000.00'
	urkro['Kode']=urkro['kdprog']+'.'+urkro['kdkro']
	urkro = urkro[['ID', 'Kode','Uraian']]
	urkro.reset_index(drop=True, inplace=True)

	urpisah = pd.concat([kdprog,idkro], axis=1)
	urpisah.loc[:,'kdprog'] = urpisah.loc[:,'kdprog'].ffill()
	urpisah=urpisah.dropna()
	urpisah['ID']=urpisah['kdprog']+'.'+urpisah['kdkro']+'.000.000.00'
	urpisah['Kode']=np.nan
	urpisah['Uraian']=np.nan
	urpisah = urpisah[['ID', 'Kode','Uraian']]
	urpisah.reset_index(drop=True, inplace=True)

	## Gabung Seluruh Data

	satker = pd.concat([urpisah,urkro,urro,rpdkompsum])
	satker.sort_values(by=['ID','Kode'],na_position='first',inplace=True,ignore_index=True)
	satker.drop(index=satker.index[0], axis=0, inplace=True)
	satker['Satker']=kode_satker

	# Tulis Excel

	## Ketikan RPD Terakhir
	satker['Indeks'] = range(1, len(satker) + 1)
	satker['Indeks'] = satker['Indeks']+19 
	satker['Indeks'] = satker['Indeks'].astype(str)

	satker['Sisa RPD'].loc[~satker['Pagu'].isnull()] = '=G'+satker['Indeks']+'-SUM(I'+satker['Indeks']+':T'+satker['Indeks']+')'

	satker = satker.reindex(['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

	satker = satker.fillna('')
	
	# Create a Pandas Excel writer using XlsxWriter as the engine.
	excelname = '/content/'+'[DJPb Babel] RPD Terakhir ' + kode_satker + '.xlsx'
	writer = pd.ExcelWriter(buffer, engine='xlsxwriter')

	# Get the xlsxwriter workbook and worksheet objects.
	workbook  = writer.book

	# Format sel.
	header_format = workbook.add_format({
		'font_size':11,
		'bold': True,
		'text_wrap': True,
		'align': 'center',
		'valign': 'vcenter',
		'fg_color': '#4c9fff',
		'border': 1})
	kode_format = workbook.add_format({
		'font_size':10,
		'align': 'right',
		'bg_color': '#ffffff',
		'border': 1})
	uraian_format = workbook.add_format({
		'font_size':10,
		'align': 'left',
		'bg_color': '#ffffff',
		'border': 1})
	belppk_format = workbook.add_format({
		'font_size':10,
		'align': 'center',
		'bg_color': '#ffffff',
		'border': 1})
	angka_format = workbook.add_format({
		'font_size':10,
		'num_format':'_(* #,##0_)',
		'bg_color': '#ffffff',
		'border': 1})
	title_format = workbook.add_format({
		'font_size':12,
		'bold': True,
		'bg_color': '#ffffff',
		'border': 0,
		'align': 'center',
		'valign': 'vcenter'})
	subtitle_format = workbook.add_format({
		'font_size':12,
		'bold': True,
		'bg_color': '#ffffff',
		'border': 0,
		'align': 'center',
		'valign': 'vcenter'})
	info_format = workbook.add_format({
		'font_size':10,
		'bg_color': '#ffffff',
		'border': 0,
		'align': 'left'})
	satker_format = workbook.add_format({
		'font_size':10,
		'bg_color': '#ffffff',
		'border': 0,
		'align': 'left',
		'valign': 'top',})
	subheader_format = workbook.add_format({
		'font_size':12,
		'bold': True,
		'bg_color': '#FCD5B4',
		'border': 1,
		'align': 'center',
		'valign': 'vcenter'})
	sumrpd_format = workbook.add_format({
		'font_size':10,
		'bold': True,
		'bg_color': '#DAEEF3',
		'border': 1,
		'align': 'right'})
	detilrpd_format = workbook.add_format({
		'font_size':10,
		'bg_color': '#ffffff',
		'border': 1,
		'align': 'right'})
	angkasum_format = workbook.add_format({
		'font_size':10,
		'num_format':'_(* #,##0_)',
		'bold': True,
		'bg_color': '#DAEEF3',
		'border': 1,
		'align': 'center'})
	angkarpd_format = workbook.add_format({
		'font_size':10,
		'num_format':'_(* #,##0_)',
		'bg_color': '#FF0000',
		'border': 1})
	beranda_format = workbook.add_format({
		'font_size':28,
		'bold': True,})
	kiri_format = workbook.add_format({
		'bold': True,})
	tengah_format = workbook.add_format({
		'align': 'center',			
		'bold': True,})
	angkaberanda_format = workbook.add_format({
		'num_format':'_(* #,##0_)',})
	
	percent_format = workbook.add_format({'num_format': '0.00%'})
	
	# Mulai mengetik dafatrame ke excel sheet RPD Terakhir
	satker.to_excel(writer, sheet_name='RPD Terakhir', startrow=19, header=False, index=False)
	worksheet = writer.sheets['RPD Terakhir']

	# Write the column headers with the defined format.
	for col_num, value in enumerate(satker.columns.values):
		worksheet.write(9, col_num, value, header_format)

	worksheet.set_zoom(90)
	worksheet.set_column('A:B', None, None, {'hidden': 1})
	worksheet.set_row(2, None, info_format)
	worksheet.set_row(3, None, info_format)
	worksheet.set_row(4, None, info_format)
	worksheet.set_row(5, None, info_format)
	worksheet.set_row(6, None, info_format)
	worksheet.set_row(7, None, info_format)
	worksheet.set_row(8, None, info_format)

	worksheet.set_column_pixels('C:C', 100, kode_format)
	worksheet.set_column_pixels('D:D', 200, uraian_format)
	worksheet.set_column_pixels('E:F', 77, belppk_format)
	worksheet.set_column_pixels('G:T', 110, angka_format)

	worksheet.merge_range('A1:T1','RENCANA PENARIKAN DANA (RPD) TERAKHIR TA ____', title_format)
	worksheet.merge_range('A2:T2','INOVASI BIDANG PPA I KANWIL DJPB BANGKA BELITUNG',subtitle_format)
	worksheet.write_string('C3', 'Satuan Kerja')
	worksheet.write_string('D3', ': '+nama_satker,satker_format)

	worksheet.write_string('C4', 'Hal Perlu Diperhatikan:')
	worksheet.write_string('C5', '1. Sesuaikan angka Pagu Komponen terlebih dahulu dengan pagu hasil revisi di SAKTI. Jika pagu berubah, sesuaikan RPD mengikuti Modul yang telah disediakan.')
	worksheet.write_string('C6', '2. Untuk Revisi Pemutakhiran POK di KPA silakan gunakan sheet ini untuk pengisian RPD Komponen yang berubah karena revisi POK.')
	worksheet.write_string('C7', '3. Untuk Pemutakhiran Hal III DIPA, update dulu angka sheet "RPD Terakhir" ini dengan data Laporan Fa Detail Basis Kas di SAKTI.')
	worksheet.write_string('C8', '4. Kolom "Sisa RPD" akan menampilkan sisa dana yang perlu direncanakan bulan berikutnya. Pastikan kolom Sisa RPD tidak merah (bernilai 0) yang menunjukkan semua pagu sudah diRPD-kan.')

	worksheet.merge_range('C11:T11', 'Informasi Target Penyerapan',subheader_format)
	worksheet.merge_range('C19:T19', 'Rencana Penarikan Dana (RPD)',subheader_format)
	worksheet.merge_range('C12:F12', 'Sisa Target Penyerapan Triwulan',sumrpd_format)
	worksheet.merge_range('C13:F13', 'Nominal Target Penyerapan Triwulan',sumrpd_format)
	worksheet.merge_range('C14:F14', 'Akumulasi Rencana Penarikan Dana Triwulan',sumrpd_format)
	worksheet.merge_range('C15:F15', 'Pegawai',detilrpd_format)
	worksheet.merge_range('C16:F16', 'Barang',detilrpd_format)
	worksheet.merge_range('C17:F17', 'Modal',detilrpd_format)
	worksheet.merge_range('C18:F18', 'Bantuan Sosial',detilrpd_format)
	worksheet.merge_range('H12:H14',None,subheader_format)

	worksheet.write_dynamic_array_formula('G12', '=G13-G14',angkasum_format)
	worksheet.merge_range('I12:K12', None)
	worksheet.merge_range('L12:N12', None)
	worksheet.merge_range('O12:Q12', None)
	worksheet.merge_range('R12:T12', None)
	worksheet.write_dynamic_array_formula('I12', '=IF(I13-I14<0,"Sudah sesuai/melebihi target triwulan",I13-I14)',angkasum_format)
	worksheet.write_dynamic_array_formula('L12', '=IF(L13-L14<0,"Sudah sesuai/melebihi target triwulan",L13-L14)',angkasum_format)
	worksheet.write_dynamic_array_formula('O12', '=IF(O13-O14<0,"Sudah sesuai/melebihi target triwulan",O13-O14)',angkasum_format)
	worksheet.write_dynamic_array_formula('R12', '=IF(R13-R14<0,"Sudah sesuai/melebihi target triwulan",R13-R14)',angkasum_format)

	worksheet.write_dynamic_array_formula('G13', '=SUM(R13)',angkasum_format)
	worksheet.merge_range('I13:K13', None)
	worksheet.merge_range('L13:N13', None)
	worksheet.merge_range('O13:Q13', None)
	worksheet.merge_range('R13:T13', None)
	worksheet.write_string('I13', 'Isi nominal target OM SPAN',angkasum_format)
	worksheet.write_string('L13', 'Isi nominal target OM SPAN',angkasum_format)
	worksheet.write_string('O13', 'Isi nominal target OM SPAN',angkasum_format)
	worksheet.write_string('R13', 'Isi nominal target OM SPAN',angkasum_format)

	worksheet.write_dynamic_array_formula('G14', '=SUM(R14)',angkasum_format)
	worksheet.merge_range('I14:K14', None)
	worksheet.merge_range('L14:N14', None)
	worksheet.merge_range('O14:Q14', None)
	worksheet.merge_range('R14:T14', None)
	worksheet.write_dynamic_array_formula('I14', '=SUM(I15:K18)',angkasum_format)
	worksheet.write_dynamic_array_formula('L14', '=SUM(I15:N18)',angkasum_format)
	worksheet.write_dynamic_array_formula('O14', '=SUM(I15:Q18)',angkasum_format)
	worksheet.write_dynamic_array_formula('R14', '=SUM(I15:T18)',angkasum_format)

	worksheet.write_dynamic_array_formula('G15', '=SUMIF($E:$E,$C15,$G:$G)')
	worksheet.write_dynamic_array_formula('H15', '=G15-SUM(I15:T15)')
	worksheet.write_dynamic_array_formula('I15', '=SUMIF($E:$E,$C15,I:I)')
	worksheet.write_dynamic_array_formula('J15', '=SUMIF($E:$E,$C15,J:J)')
	worksheet.write_dynamic_array_formula('K15', '=SUMIF($E:$E,$C15,K:K)')
	worksheet.write_dynamic_array_formula('L15', '=SUMIF($E:$E,$C15,L:L)')
	worksheet.write_dynamic_array_formula('M15', '=SUMIF($E:$E,$C15,M:M)')
	worksheet.write_dynamic_array_formula('N15', '=SUMIF($E:$E,$C15,N:N)')
	worksheet.write_dynamic_array_formula('O15', '=SUMIF($E:$E,$C15,O:O)')
	worksheet.write_dynamic_array_formula('P15', '=SUMIF($E:$E,$C15,P:P)')
	worksheet.write_dynamic_array_formula('Q15', '=SUMIF($E:$E,$C15,Q:Q)')
	worksheet.write_dynamic_array_formula('R15', '=SUMIF($E:$E,$C15,R:R)')
	worksheet.write_dynamic_array_formula('S15', '=SUMIF($E:$E,$C15,S:S)')
	worksheet.write_dynamic_array_formula('T15', '=SUMIF($E:$E,$C15,T:T)')

	worksheet.write_dynamic_array_formula('G16', '=SUMIF($E:$E,$C16,$G:$G)')
	worksheet.write_dynamic_array_formula('H16', '=G16-SUM(I16:T16)')
	worksheet.write_dynamic_array_formula('I16', '=SUMIF($E:$E,$C16,I:I)')
	worksheet.write_dynamic_array_formula('J16', '=SUMIF($E:$E,$C16,J:J)')
	worksheet.write_dynamic_array_formula('K16', '=SUMIF($E:$E,$C16,K:K)')
	worksheet.write_dynamic_array_formula('L16', '=SUMIF($E:$E,$C16,L:L)')
	worksheet.write_dynamic_array_formula('M16', '=SUMIF($E:$E,$C16,M:M)')
	worksheet.write_dynamic_array_formula('N16', '=SUMIF($E:$E,$C16,N:N)')
	worksheet.write_dynamic_array_formula('O16', '=SUMIF($E:$E,$C16,O:O)')
	worksheet.write_dynamic_array_formula('P16', '=SUMIF($E:$E,$C16,P:P)')
	worksheet.write_dynamic_array_formula('Q16', '=SUMIF($E:$E,$C16,Q:Q)')
	worksheet.write_dynamic_array_formula('R16', '=SUMIF($E:$E,$C16,R:R)')
	worksheet.write_dynamic_array_formula('S16', '=SUMIF($E:$E,$C16,S:S)')
	worksheet.write_dynamic_array_formula('T16', '=SUMIF($E:$E,$C16,T:T)')

	worksheet.write_dynamic_array_formula('G17', '=SUMIF($E:$E,$C17,$G:$G)')
	worksheet.write_dynamic_array_formula('H17', '=G17-SUM(I17:T17)')
	worksheet.write_dynamic_array_formula('I17', '=SUMIF($E:$E,$C17,I:I)')
	worksheet.write_dynamic_array_formula('J17', '=SUMIF($E:$E,$C17,J:J)')
	worksheet.write_dynamic_array_formula('K17', '=SUMIF($E:$E,$C17,K:K)')
	worksheet.write_dynamic_array_formula('L17', '=SUMIF($E:$E,$C17,L:L)')
	worksheet.write_dynamic_array_formula('M17', '=SUMIF($E:$E,$C17,M:M)')
	worksheet.write_dynamic_array_formula('N17', '=SUMIF($E:$E,$C17,N:N)')
	worksheet.write_dynamic_array_formula('O17', '=SUMIF($E:$E,$C17,O:O)')
	worksheet.write_dynamic_array_formula('P17', '=SUMIF($E:$E,$C17,P:P)')
	worksheet.write_dynamic_array_formula('Q17', '=SUMIF($E:$E,$C17,Q:Q)')
	worksheet.write_dynamic_array_formula('R17', '=SUMIF($E:$E,$C17,R:R)')
	worksheet.write_dynamic_array_formula('S17', '=SUMIF($E:$E,$C17,S:S)')
	worksheet.write_dynamic_array_formula('T17', '=SUMIF($E:$E,$C17,T:T)')

	worksheet.write_dynamic_array_formula('G18', '=SUMIF($E:$E,$C18,$G:$G)')
	worksheet.write_dynamic_array_formula('H18', '=G18-SUM(I18:T18)')
	worksheet.write_dynamic_array_formula('I18', '=SUMIF($E:$E,$C18,I:I)')
	worksheet.write_dynamic_array_formula('J18', '=SUMIF($E:$E,$C18,J:J)')
	worksheet.write_dynamic_array_formula('K18', '=SUMIF($E:$E,$C18,K:K)')
	worksheet.write_dynamic_array_formula('L18', '=SUMIF($E:$E,$C18,L:L)')
	worksheet.write_dynamic_array_formula('M18', '=SUMIF($E:$E,$C18,M:M)')
	worksheet.write_dynamic_array_formula('N18', '=SUMIF($E:$E,$C18,N:N)')
	worksheet.write_dynamic_array_formula('O18', '=SUMIF($E:$E,$C18,O:O)')
	worksheet.write_dynamic_array_formula('P18', '=SUMIF($E:$E,$C18,P:P)')
	worksheet.write_dynamic_array_formula('Q18', '=SUMIF($E:$E,$C18,Q:Q)')
	worksheet.write_dynamic_array_formula('R18', '=SUMIF($E:$E,$C18,R:R)')
	worksheet.write_dynamic_array_formula('S18', '=SUMIF($E:$E,$C18,S:S)')
	worksheet.write_dynamic_array_formula('T18', '=SUMIF($E:$E,$C18,T:T)')

	worksheet.conditional_format('H15:H1048576', {'type': 'cell',
											 'criteria': '<>',
											 'value': 0,
											 'format': angkarpd_format})

	worksheet.ignore_errors({'number_stored_as_text': 'C20:C1048576'})

	worksheet.freeze_panes('I11')

	workbook.set_properties({
		'title':    'Perawas RPD',
		'subject':  'Perencanaan dan Pengawasan RPD',
		'author':   'Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung',
		'company':  'Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung',
		'category': 'Perencanaan Kas',
		'keywords': 'Perencanaan, Kas, Keuangan',
		'comments': 'Inovasi dari Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung'})

	writer.save()

	excelname = '[DJPb Babel] RPD Terakhir Satker ' + kode_satker + '.xlsx'
	
	st.download_button(
		label="Download RPD Terakhir Satker Anda di sini",
		data=buffer,
		file_name=excelname,
		mime="application/vnd.ms-excel"
	)
	
	st.markdown('---')
	st.write("Jika ingin mengunduh data realisasi, lanjut upload file MonSAKTI di bawah.")

	uploaded_file2 = st.file_uploader('Upload MONSAKTI Fa Detail di sini. Nama file unduhan "Monitoring Detail Transaksi...".', type='xlsx')
	if uploaded_file2:
		st.markdown('---')
		real = pd.read_excel(uploaded_file2, index_col=None, header=2, engine='openpyxl')
		real.reset_index(drop=True, inplace=True)
		real2 = real[["TANGGAL SP2D", "KODE COA",	"NILAI RUPIAH"]]
		real2.rename(columns={'NILAI RUPIAH': 'nilai'}, inplace=True)
		real2[['satker', 'kppn', 'Akun','program','kro','sdana','bank','kewenangan','lokasi','budget','xxx','xxxx','ro','Komponen','Sub Komponen','xxxxx']] = real2['KODE COA'].str.split('.', expand=True)
		real2['ID'] = real2['program'].str[-2:]+"."+real2['kro'].str[:4]+"."+real2['kro'].str[-3:]+"."+real2['ro']+"."+real2['Komponen']+"."+real2['Akun'].str[:2]
		real2['TANGGAL SP2D']=pd.to_datetime(real2['TANGGAL SP2D'])
		real2['bulan'] = real2['TANGGAL SP2D'].dt.month_name().str[:3]
		real2['Jenis Belanja'] = real2['Akun'].str[:2]
		real2 = real2[["ID", "bulan",	"nilai"]]
		real2=pd.pivot_table(real2, values='nilai', index='ID', columns='bulan', aggfunc='sum', fill_value=0, dropna=True, sort=True)
		real2.reset_index(inplace=True)
		idreal = rpdkompsum[['ID']]
		real13 = pd.merge(idreal,real2,on='ID',how='left')
		real4 = real13.fillna(0)
		
		real4 = real4.reindex(['ID','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

		## Ketikan Realisasi
		rpdreal = satker[['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD',]]
		rpdreal1 = pd.merge(rpdreal,real4,on='ID',how='left')

		# Create a Pandas Excel writer using XlsxWriter as the engine.
		excelname = '/content/'+'[DJPb Babel] RPD Realisasi Satker ' + kode_satker + '.xlsx'
		writer = pd.ExcelWriter(buffer2, engine='xlsxwriter')

		# Get the xlsxwriter workbook and worksheet objects.
		workbook  = writer.book


		# Format sel.
		header_format = workbook.add_format({
			'font_size':11,
			'bold': True,
			'text_wrap': True,
			'align': 'center',
			'valign': 'vcenter',
			'fg_color': '#4c9fff',
			'border': 1})
		kode_format = workbook.add_format({
			'font_size':10,
			'align': 'right',
			'bg_color': '#ffffff',
			'border': 1})
		uraian_format = workbook.add_format({
			'font_size':10,
			'align': 'left',
			'bg_color': '#ffffff',
			'border': 1})
		belppk_format = workbook.add_format({
			'font_size':10,
			'align': 'center',
			'bg_color': '#ffffff',
			'border': 1})
		angka_format = workbook.add_format({
			'font_size':10,
			'num_format':'_(* #,##0_)',
			'bg_color': '#ffffff',
			'border': 1})
		title_format = workbook.add_format({
			'font_size':12,
			'bold': True,
			'bg_color': '#ffffff',
			'border': 0,
			'align': 'center',
			'valign': 'vcenter'})
		subtitle_format = workbook.add_format({
			'font_size':12,
			'bold': True,
			'bg_color': '#ffffff',
			'border': 0,
			'align': 'center',
			'valign': 'vcenter'})
		info_format = workbook.add_format({
			'font_size':10,
			'bg_color': '#ffffff',
			'border': 0,
			'align': 'left'})
		satker_format = workbook.add_format({
			'font_size':10,
			'bg_color': '#ffffff',
			'border': 0,
			'align': 'left',
			'valign': 'top',})
		subheader_format = workbook.add_format({
			'font_size':12,
			'bold': True,
			'bg_color': '#FCD5B4',
			'border': 1,
			'align': 'center',
			'valign': 'vcenter'})
		sumrpd_format = workbook.add_format({
			'font_size':10,
			'bold': True,
			'bg_color': '#DAEEF3',
			'border': 1,
			'align': 'right'})
		detilrpd_format = workbook.add_format({
			'font_size':10,
			'bg_color': '#ffffff',
			'border': 1,
			'align': 'right'})
		angkasum_format = workbook.add_format({
			'font_size':10,
			'num_format':'_(* #,##0_)',
			'bold': True,
			'bg_color': '#DAEEF3',
			'border': 1,
			'align': 'center'})
		angkarpd_format = workbook.add_format({
			'font_size':10,
			'num_format':'_(* #,##0_)',
			'bg_color': '#FF0000',
			'border': 1})
		beranda_format = workbook.add_format({
			'font_size':28,
			'bold': True,})
		kiri_format = workbook.add_format({
			'bold': True,})
		tengah_format = workbook.add_format({
			'align': 'center',			
			'bold': True,})
		angkaberanda_format = workbook.add_format({
			'num_format':'_(* #,##0_)',})


		# Mulai mengetik dafatrame ke excel sheet RPD Realisasi
		rpdreal1.to_excel(writer, sheet_name='RPD Realisasi', startrow=19, header=False, index=False)
		worksheet = writer.sheets['RPD Realisasi']

		# Write the column headers with the defined format.
		for col_num, value in enumerate(satker.columns.values):
			worksheet.write(9, col_num, value, header_format)

		worksheet.set_zoom(90)
		worksheet.set_column('A:B', None, None, {'hidden': 1})
		worksheet.set_row(2, None, info_format)
		worksheet.set_row(3, None, info_format)
		worksheet.set_row(4, None, info_format)
		worksheet.set_row(5, None, info_format)
		worksheet.set_row(6, None, info_format)
		worksheet.set_row(7, None, info_format)
		worksheet.set_row(8, None, info_format)

		worksheet.set_column_pixels('C:C', 100, kode_format)
		worksheet.set_column_pixels('D:D', 200, uraian_format)
		worksheet.set_column_pixels('E:F', 77, belppk_format)
		worksheet.set_column_pixels('G:T', 110, angka_format)

		worksheet.merge_range('A1:T1','RENCANA PENARIKAN DANA (RPD) REALISASI TA ____', title_format)
		worksheet.merge_range('A2:T2','INOVASI BIDANG PPA I KANWIL DJPB BANGKA BELITUNG',subtitle_format)
		worksheet.write_string('C3', 'Satuan Kerja')
		worksheet.write_string('D3', ': '+nama_satker,satker_format)

		worksheet.write_string('C4', 'Hal Perlu Diperhatikan:')
		worksheet.write_string('C5', '1. Sesuaikan angka Pagu Komponen terlebih dahulu dengan pagu hasil revisi di SAKTI. Jika pagu berubah, sesuaikan RPD mengikuti Modul yang telah disediakan.')
		worksheet.write_string('C6', '2. Untuk Revisi Pemutakhiran POK di KPA silakan gunakan sheet ini untuk pengisian RPD Komponen yang berubah karena revisi POK.')
		worksheet.write_string('C7', '3. Untuk Pemutakhiran Hal III DIPA, update dulu angka sheet "RPD Terakhir" ini dengan data Laporan Fa Detail Basis Kas di SAKTI.')
		worksheet.write_string('C8', '4. Kolom "Sisa RPD" akan menampilkan sisa dana yang perlu direncanakan bulan berikutnya. Pastikan kolom Sisa RPD tidak merah (bernilai 0) yang menunjukkan semua pagu sudah diRPD-kan.')

		worksheet.merge_range('C11:T11', 'Informasi Target Penyerapan',subheader_format)
		worksheet.merge_range('C19:T19', 'Rencana Penarikan Dana (RPD)',subheader_format)
		worksheet.merge_range('C12:F12', 'Sisa Target Penyerapan Triwulan',sumrpd_format)
		worksheet.merge_range('C13:F13', 'Nominal Target Penyerapan Triwulan',sumrpd_format)
		worksheet.merge_range('C14:F14', 'Akumulasi Rencana Penarikan Dana Triwulan',sumrpd_format)
		worksheet.merge_range('C15:F15', 'Pegawai',detilrpd_format)
		worksheet.merge_range('C16:F16', 'Barang',detilrpd_format)
		worksheet.merge_range('C17:F17', 'Modal',detilrpd_format)
		worksheet.merge_range('C18:F18', 'Bantuan Sosial',detilrpd_format)
		worksheet.merge_range('H12:H14',None,subheader_format)

		worksheet.write_dynamic_array_formula('G12', '=G13-G14',angkasum_format)
		worksheet.merge_range('I12:K12', None)
		worksheet.merge_range('L12:N12', None)
		worksheet.merge_range('O12:Q12', None)
		worksheet.merge_range('R12:T12', None)
		worksheet.write_dynamic_array_formula('I12', '=IF(I13-I14<0,"Sudah sesuai/melebihi target triwulan",I13-I14)',angkasum_format)
		worksheet.write_dynamic_array_formula('L12', '=IF(L13-L14<0,"Sudah sesuai/melebihi target triwulan",L13-L14)',angkasum_format)
		worksheet.write_dynamic_array_formula('O12', '=IF(O13-O14<0,"Sudah sesuai/melebihi target triwulan",O13-O14)',angkasum_format)
		worksheet.write_dynamic_array_formula('R12', '=IF(R13-R14<0,"Sudah sesuai/melebihi target triwulan",R13-R14)',angkasum_format)

		worksheet.write_dynamic_array_formula('G13', '=SUM(R13)',angkasum_format)
		worksheet.merge_range('I13:K13', None)
		worksheet.merge_range('L13:N13', None)
		worksheet.merge_range('O13:Q13', None)
		worksheet.merge_range('R13:T13', None)
		worksheet.write_string('I13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet.write_string('L13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet.write_string('O13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet.write_string('R13', 'Isi nominal target OM SPAN',angkasum_format)

		worksheet.write_dynamic_array_formula('G14', '=SUM(R14)',angkasum_format)
		worksheet.merge_range('I14:K14', None)
		worksheet.merge_range('L14:N14', None)
		worksheet.merge_range('O14:Q14', None)
		worksheet.merge_range('R14:T14', None)
		worksheet.write_dynamic_array_formula('I14', '=SUM(I15:K18)',angkasum_format)
		worksheet.write_dynamic_array_formula('L14', '=SUM(I15:N18)',angkasum_format)
		worksheet.write_dynamic_array_formula('O14', '=SUM(I15:Q18)',angkasum_format)
		worksheet.write_dynamic_array_formula('R14', '=SUM(I15:T18)',angkasum_format)

		worksheet.write_dynamic_array_formula('G15', '=SUMIF($E:$E,$C15,$G:$G)')
		worksheet.write_dynamic_array_formula('H15', '=G15-SUM(I15:T15)')
		worksheet.write_dynamic_array_formula('I15', '=SUMIF($E:$E,$C15,I:I)')
		worksheet.write_dynamic_array_formula('J15', '=SUMIF($E:$E,$C15,J:J)')
		worksheet.write_dynamic_array_formula('K15', '=SUMIF($E:$E,$C15,K:K)')
		worksheet.write_dynamic_array_formula('L15', '=SUMIF($E:$E,$C15,L:L)')
		worksheet.write_dynamic_array_formula('M15', '=SUMIF($E:$E,$C15,M:M)')
		worksheet.write_dynamic_array_formula('N15', '=SUMIF($E:$E,$C15,N:N)')
		worksheet.write_dynamic_array_formula('O15', '=SUMIF($E:$E,$C15,O:O)')
		worksheet.write_dynamic_array_formula('P15', '=SUMIF($E:$E,$C15,P:P)')
		worksheet.write_dynamic_array_formula('Q15', '=SUMIF($E:$E,$C15,Q:Q)')
		worksheet.write_dynamic_array_formula('R15', '=SUMIF($E:$E,$C15,R:R)')
		worksheet.write_dynamic_array_formula('S15', '=SUMIF($E:$E,$C15,S:S)')
		worksheet.write_dynamic_array_formula('T15', '=SUMIF($E:$E,$C15,T:T)')

		worksheet.write_dynamic_array_formula('G16', '=SUMIF($E:$E,$C16,$G:$G)')
		worksheet.write_dynamic_array_formula('H16', '=G16-SUM(I16:T16)')
		worksheet.write_dynamic_array_formula('I16', '=SUMIF($E:$E,$C16,I:I)')
		worksheet.write_dynamic_array_formula('J16', '=SUMIF($E:$E,$C16,J:J)')
		worksheet.write_dynamic_array_formula('K16', '=SUMIF($E:$E,$C16,K:K)')
		worksheet.write_dynamic_array_formula('L16', '=SUMIF($E:$E,$C16,L:L)')
		worksheet.write_dynamic_array_formula('M16', '=SUMIF($E:$E,$C16,M:M)')
		worksheet.write_dynamic_array_formula('N16', '=SUMIF($E:$E,$C16,N:N)')
		worksheet.write_dynamic_array_formula('O16', '=SUMIF($E:$E,$C16,O:O)')
		worksheet.write_dynamic_array_formula('P16', '=SUMIF($E:$E,$C16,P:P)')
		worksheet.write_dynamic_array_formula('Q16', '=SUMIF($E:$E,$C16,Q:Q)')
		worksheet.write_dynamic_array_formula('R16', '=SUMIF($E:$E,$C16,R:R)')
		worksheet.write_dynamic_array_formula('S16', '=SUMIF($E:$E,$C16,S:S)')
		worksheet.write_dynamic_array_formula('T16', '=SUMIF($E:$E,$C16,T:T)')

		worksheet.write_dynamic_array_formula('G17', '=SUMIF($E:$E,$C17,$G:$G)')
		worksheet.write_dynamic_array_formula('H17', '=G17-SUM(I17:T17)')
		worksheet.write_dynamic_array_formula('I17', '=SUMIF($E:$E,$C17,I:I)')
		worksheet.write_dynamic_array_formula('J17', '=SUMIF($E:$E,$C17,J:J)')
		worksheet.write_dynamic_array_formula('K17', '=SUMIF($E:$E,$C17,K:K)')
		worksheet.write_dynamic_array_formula('L17', '=SUMIF($E:$E,$C17,L:L)')
		worksheet.write_dynamic_array_formula('M17', '=SUMIF($E:$E,$C17,M:M)')
		worksheet.write_dynamic_array_formula('N17', '=SUMIF($E:$E,$C17,N:N)')
		worksheet.write_dynamic_array_formula('O17', '=SUMIF($E:$E,$C17,O:O)')
		worksheet.write_dynamic_array_formula('P17', '=SUMIF($E:$E,$C17,P:P)')
		worksheet.write_dynamic_array_formula('Q17', '=SUMIF($E:$E,$C17,Q:Q)')
		worksheet.write_dynamic_array_formula('R17', '=SUMIF($E:$E,$C17,R:R)')
		worksheet.write_dynamic_array_formula('S17', '=SUMIF($E:$E,$C17,S:S)')
		worksheet.write_dynamic_array_formula('T17', '=SUMIF($E:$E,$C17,T:T)')

		worksheet.write_dynamic_array_formula('G18', '=SUMIF($E:$E,$C18,$G:$G)')
		worksheet.write_dynamic_array_formula('H18', '=G18-SUM(I18:T18)')
		worksheet.write_dynamic_array_formula('I18', '=SUMIF($E:$E,$C18,I:I)')
		worksheet.write_dynamic_array_formula('J18', '=SUMIF($E:$E,$C18,J:J)')
		worksheet.write_dynamic_array_formula('K18', '=SUMIF($E:$E,$C18,K:K)')
		worksheet.write_dynamic_array_formula('L18', '=SUMIF($E:$E,$C18,L:L)')
		worksheet.write_dynamic_array_formula('M18', '=SUMIF($E:$E,$C18,M:M)')
		worksheet.write_dynamic_array_formula('N18', '=SUMIF($E:$E,$C18,N:N)')
		worksheet.write_dynamic_array_formula('O18', '=SUMIF($E:$E,$C18,O:O)')
		worksheet.write_dynamic_array_formula('P18', '=SUMIF($E:$E,$C18,P:P)')
		worksheet.write_dynamic_array_formula('Q18', '=SUMIF($E:$E,$C18,Q:Q)')
		worksheet.write_dynamic_array_formula('R18', '=SUMIF($E:$E,$C18,R:R)')
		worksheet.write_dynamic_array_formula('S18', '=SUMIF($E:$E,$C18,S:S)')
		worksheet.write_dynamic_array_formula('T18', '=SUMIF($E:$E,$C18,T:T)')

		worksheet.conditional_format('H15:H1048576', {'type': 'cell',
												 'criteria': '<>',
												 'value': 0,
												 'format': angkarpd_format})

		worksheet.ignore_errors({'number_stored_as_text': 'C20:C1048576'})

		worksheet.freeze_panes('I11')

		workbook.set_properties({
			'title':    'Perawas RPD',
			'subject':  'Perencanaan dan Pengawasan RPD',
			'author':   'Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung',
			'company':  'Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung',
			'category': 'Perencanaan Kas',
			'keywords': 'Perencanaan, Kas, Keuangan',
			'comments': 'Inovasi dari Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung'})

		writer.save()

		excelname = '[DJPb Babel] RPD Realisasi Satker ' + kode_satker + '.xlsx'
		
		st.download_button(
			label="Download RPD Realisasi Satker Anda di sini",
			data=buffer2,
			file_name=excelname,
			mime="application/vnd.ms-excel"
		)
	
st.markdown('---')
st.caption('Created by Farhan Ariq R. - Bidang PPA I, Kanwil DJPb Bangka Belitung 2022')
