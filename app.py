import streamlit as st
import pandas as pd
import io
import numpy as np
from csv import writer
import csv
import datetime
from datetime import datetime
import warnings
import gspread

from pandas.errors import SettingWithCopyWarning
warnings.simplefilter(action="ignore", category=SettingWithCopyWarning)

# Untuk Download
buffer = io.BytesIO()
buffer2 = io.BytesIO()

# Halaman Streamlit
st.set_page_config(page_title='[DJPb Babel] Perawas RPD')
st.title('Inovasi Perawas RPD 📋')
st.subheader('Dari Kanwil DJPb Babel untuk Indonesia ❤️')
st.write("Aplikasi Perencanaan dan Pengawasan RPD (Perawas RPD) berfungsi membantu proses penyusunan RPD Satker Anda.")
st.write("Follow [Instagram](https://www.instagram.com/djpbbabel/) dan [Youtube](https://www.youtube.com/channel/UCM70tgByAEPpIPqME2mKKyQ).")
st.write("Panduan aplikasi [klik di sini](http://bit.ly/materidjpbbabel).")
st.markdown('---')

tab1, tab2 = st.tabs(["Revisi Pemutakhiran KPA","Revisi Halaman III DIPA"])

with tab1:
	st.header("Unduh RPD DIPA Terakhir untuk membantu Revisi Pemutakhiran KPA")
	# Upload File RPD DIPA Usulan
	uploaded_file = st.file_uploader('Upload File RPD DIPA Usulan di sini.', type='xlsx')
	uploaded_file2 = st.file_uploader('Upload File RPD DIPA Petikan Terakhir di sini.', type='xlsx')
	
	if uploaded_file and uploaded_file2:
		raw = pd.read_excel(uploaded_file2, index_col=None, header=6, skipfooter=5, engine='openpyxl')
		info = pd.read_excel(uploaded_file2, index_col=None, nrows = 1, dtype=str, engine='openpyxl')
		info.rename(columns = {'Unnamed: 4':'kdsatker', 'Unnamed: 6':'nmsatker'}, inplace = True)
		info['kdsatker'] = info['kdsatker'].str[-6:]
		nama_satker = info['nmsatker'].iloc[0]  +' (' + info['kdsatker'].iloc[0] + ')'
		kode_satker = info['kdsatker'].iloc[0]

		# Mencatat log pengguna
		timestamp_now = datetime.now()
		stringTimestamp = timestamp_now.strftime("%Y-%m-%d %H:%M:%S")
		setLog = [stringTimestamp, "Revisi KPA",kodeSatker, namaSatker]

		# Mengautentikasi dengan kunci API
		gc = gspread.service_account(st.secrets["gs_service_account"])

		# Buka spreadsheet
		spreadsheet = gc.open("Log User Perawas RPD").sheet1

		# Menambahkan data ke dalam baris baru
		spreadsheet.append_row(setLog)

		# Master Data RPD

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
		idskomp1 = idmaster.loc[idmaster['Kode'].str.len() == 1]
		idskomp2 = idmaster.loc[idmaster['Kode'].str.len() == 2]
		idskomp1['Kode']='0'+idskomp1['Kode']
		idskomp = pd.concat([idskomp1,idskomp2], axis=0)
		idakun = idmaster.loc[idmaster['Kode'].str.len() == 6]

		# Rename Nama Kolom id

		idprog.columns = ['kdprog', 'Uraian']
		idgiat.columns = ['kdgiat', 'Uraian']
		idkro.columns = ['kdkro', 'Uraian']
		idro.columns = ['kdro', 'Uraian']
		idkomp.columns = ['kdkomp', 'Uraian']
		idskomp.columns = ['kdskomp', 'Uraian']

		# Mengambil jenis belanja saja

		idjenbel = idakun['Kode'].str[:2]
		idakun.columns = ['kdakun', 'Uraian'] #Harus rename di sini karena akan berpengaruh ke kode di bawah kalau sebelum syntax di atas

		# Penyesuaian Kode Program, KRO, dan RO

		kdprog = idprog['kdprog'].str[-2:]
		kdkro = idkro['kdkro'].str[-3:]
		kdro = idro['kdro'].str[-3:]
		kdgiat = idgiat['kdgiat']
		kdkomp = idkomp['kdkomp']
		kdskomp = idskomp['kdskomp']

		# Proses Bentuk Uraian RO & KRO

		urpisah = pd.concat([kdprog,idkro], axis=1, sort=False).sort_index()
		urpisah.loc[:,'kdprog'] = urpisah.loc[:,'kdprog'].ffill()
		urpisah=urpisah.dropna()
		urpisah['ID']=urpisah['kdprog']+'.'+urpisah['kdkro']+'.000.000.00'
		urpisah['Kode']=np.nan
		urpisah['Uraian']=np.nan
		urpisah = urpisah[['ID', 'Kode','Uraian']]
		urpisah.reset_index(drop=True, inplace=True)

		urkro = pd.concat([kdprog,idkro], axis=1, sort=False).sort_index()
		urkro.loc[:,'kdprog'] = urkro.loc[:,'kdprog'].ffill()
		urkro=urkro.dropna()
		urkro['ID']=urkro['kdprog']+'.'+urkro['kdkro']+'.000.000.00'
		urkro['Kode']=urkro['kdprog']+'.'+urkro['kdkro']
		urkro = urkro[['ID', 'Kode','Uraian']]
		urkro.reset_index(drop=True, inplace=True)

		urro = pd.concat([kdprog,kdgiat,idro], axis=1, sort=False).sort_index()
		urro.loc[:,'kdprog'] = urro.loc[:,'kdprog'].ffill()
		urro.loc[:,'kdgiat'] = urro.loc[:,'kdgiat'].ffill()
		urro=urro.dropna()
		urro['ID']=urro['kdprog']+'.'+urro['kdgiat']+'.'+urro['kdro']+'.000.00'
		urro['Kode']=urro['kdgiat']+'.'+urro['kdro']
		urro = urro[['ID', 'Kode','Uraian']]
		urro.reset_index(drop=True, inplace=True)

		urkomp = pd.concat([kdprog,kdgiat,kdkro,kdro,idkomp], axis=1, sort=False).sort_index()
		urkomp.loc[:,'kdprog'] = urkomp.loc[:,'kdprog'].ffill()
		urkomp.loc[:,'kdgiat'] = urkomp.loc[:,'kdgiat'].ffill()
		urkomp.loc[:,'kdkro'] = urkomp.loc[:,'kdkro'].ffill()
		urkomp.loc[:,'kdro'] = urkomp.loc[:,'kdro'].ffill()
		urkomp=urkomp.dropna()
		urkomp['ID']=urkomp['kdprog']+'.'+urkomp['kdgiat']+'.'+urkomp['kdkro']+'.'+urkomp['kdro']+'.'+urkomp['kdkomp']+'.00'
		urkomp['Kode']=urkomp['kdkomp']
		urkomp = urkomp[['ID', 'Kode','Uraian']]
		urkomp.reset_index(drop=True, inplace=True)

		urskomp = pd.concat([kdprog,kdgiat,kdkro,kdro,kdkomp,idskomp], axis=1, sort=False).sort_index()
		urskomp.loc[:,'kdprog'] = urskomp.loc[:,'kdprog'].ffill()
		urskomp.loc[:,'kdgiat'] = urskomp.loc[:,'kdgiat'].ffill()
		urskomp.loc[:,'kdkro'] = urskomp.loc[:,'kdkro'].ffill()
		urskomp.loc[:,'kdro'] = urskomp.loc[:,'kdro'].ffill()
		urskomp.loc[:,'kdkomp'] = urskomp.loc[:,'kdkomp'].ffill()
		urskomp=urskomp.dropna()
		urskomp['ID']=urskomp['kdprog']+'.'+urskomp['kdgiat']+'.'+urskomp['kdkro']+'.'+urskomp['kdro']+'.'+urskomp['kdkomp']+'.'+urskomp['kdskomp']+'.00'
		urskomp['Kode']=urskomp['kdskomp']
		urskomp = urskomp[['ID', 'Kode','Uraian']]
		urskomp.reset_index(drop=True, inplace=True)

		# Proses Bentuk Nilai Komponen & Jenis Belanja

		kompbel1 = pd.concat([idkomp,idjenbel], axis=1, sort=False).sort_index()

		kompbel1.loc[:,'kdkomp'] = kompbel1.loc[:,'kdkomp'].ffill()
		kompbel1.loc[:,'Uraian'] = kompbel1.loc[:,'Uraian'].ffill()
		kompbel1=kompbel1.dropna()
		kompbel1['ID']=kompbel1['kdkomp']+'.'+kompbel1['Kode']
		kompbel2 = kompbel1[['ID', 'Uraian', 'Kode']]
		kompbel2.rename(columns = {'Kode':'Belanja'}, inplace = True)

		kompbel3 = pd.concat([kompbel2,kdro], axis=1, sort=False).sort_index()
		kompbel3.loc[:,'kdro'] = kompbel3.loc[:,'kdro'].ffill()
		kompbel3=kompbel3.dropna()
		kompbel3['ID']=kompbel3['kdro']+'.'+kompbel3['ID']
		kompbel4 = kompbel3[['ID', 'Uraian', 'Belanja']]

		kompbel5 = pd.concat([kompbel4,kdkro], axis=1, sort=False).sort_index()
		kompbel5.loc[:,'kdkro'] = kompbel5.loc[:,'kdkro'].ffill()
		kompbel5=kompbel5.dropna()
		kompbel5['ID']=kompbel5['kdkro']+'.'+kompbel5['ID']
		kompbel6 = kompbel5[['ID', 'Uraian', 'Belanja']]

		kompbel7 = pd.concat([kompbel6,kdgiat], axis=1, sort=False).sort_index()
		kompbel7.loc[:,'kdgiat'] = kompbel7.loc[:,'kdgiat'].ffill()
		kompbel7=kompbel7.dropna()
		kompbel7['ID']=kompbel7['kdgiat']+'.'+kompbel7['ID']
		kompbel8 = kompbel7[['ID', 'Uraian', 'Belanja']]

		kompbel9 = pd.concat([kompbel8,kdprog], axis=1, sort=False).sort_index()
		kompbel9.loc[:,'kdprog'] = kompbel9.loc[:,'kdprog'].ffill()
		kompbel9=kompbel9.dropna()
		kompbel9['ID']=kompbel9['kdprog']+'.'+kompbel9['ID']
		kompbel10 = kompbel9[['ID', 'Uraian', 'Belanja']]

		# Proses Bentuk Nilai Subkomponen & Jenis Belanja

		kompbels1 = pd.concat([idskomp,idjenbel], axis=1, sort=False).sort_index()

		kompbels1.loc[:,'kdskomp'] = kompbels1.loc[:,'kdskomp'].ffill()
		kompbels1.loc[:,'Uraian'] = kompbels1.loc[:,'Uraian'].ffill()
		kompbels1=kompbels1.dropna()
		kompbels1['ID']=kompbels1['kdskomp']+'.'+kompbels1['Kode']
		kompbels2 = kompbels1[['ID', 'Uraian', 'Kode']]
		kompbels2.rename(columns = {'Kode':'Belanja'}, inplace = True)

		kompbels2A = pd.concat([kompbels2,kdkomp], axis=1, sort=False).sort_index()
		kompbels2A.loc[:,'kdkomp'] = kompbels2A.loc[:,'kdkomp'].ffill()
		kompbels2A=kompbels2A.dropna()
		kompbels2A['ID']=kompbels2A['kdkomp']+'.'+kompbels2A['ID']
		kompbels2B = kompbels2A[['ID', 'Uraian', 'Belanja']]

		kompbels3 = pd.concat([kompbels2B,kdro], axis=1, sort=False).sort_index()
		kompbels3.loc[:,'kdro'] = kompbels3.loc[:,'kdro'].ffill()
		kompbels3=kompbels3.dropna()
		kompbels3['ID']=kompbels3['kdro']+'.'+kompbels3['ID']
		kompbels4 = kompbels3[['ID', 'Uraian', 'Belanja']]

		kompbels5 = pd.concat([kompbels4,kdkro], axis=1, sort=False).sort_index()
		kompbels5.loc[:,'kdkro'] = kompbels5.loc[:,'kdkro'].ffill()
		kompbels5=kompbels5.dropna()
		kompbels5['ID']=kompbels5['kdkro']+'.'+kompbels5['ID']
		kompbels6 = kompbels5[['ID', 'Uraian', 'Belanja']]

		kompbels7 = pd.concat([kompbels6,kdgiat], axis=1, sort=False).sort_index()
		kompbels7.loc[:,'kdgiat'] = kompbels7.loc[:,'kdgiat'].ffill()
		kompbels7=kompbels7.dropna()
		kompbels7['ID']=kompbels7['kdgiat']+'.'+kompbels7['ID']
		kompbels8 = kompbels7[['ID', 'Uraian', 'Belanja']]

		kompbels9 = pd.concat([kompbels8,kdprog], axis=1, sort=False).sort_index()
		kompbels9.loc[:,'kdprog'] = kompbels9.loc[:,'kdprog'].ffill()
		kompbels9=kompbels9.dropna()
		kompbels9['ID']=kompbels9['kdprog']+'.'+kompbels9['ID']
		kompbels10 = kompbels9[['ID', 'Uraian', 'Belanja']]

		# Proses Bentuk Nilai Akun & Jenis Belanja

		kompbelss1 = pd.concat([idakun,idjenbel], axis=1, sort=False).sort_index()

		kompbelss1.loc[:,'kdakun'] = kompbelss1.loc[:,'kdakun'].ffill()
		kompbelss1.loc[:,'Uraian'] = kompbelss1.loc[:,'Uraian'].ffill()
		kompbelss1=kompbelss1.dropna()
		kompbelss1['ID']=kompbelss1['kdakun']+'.'+kompbelss1['Kode']
		kompbelss2 = kompbelss1[['ID', 'Uraian', 'Kode']]
		kompbelss2.rename(columns = {'Kode':'Belanja'}, inplace = True)

		kompbelss2A = pd.concat([kompbelss2,kdskomp], axis=1, sort=False).sort_index()
		kompbelss2A.loc[:,'kdskomp'] = kompbelss2A.loc[:,'kdskomp'].ffill()
		kompbelss2A=kompbelss2A.dropna()
		kompbelss2A['ID']=kompbelss2A['kdskomp']+'.'+kompbelss2A['ID']
		kompbelss2B = kompbelss2A[['ID', 'Uraian', 'Belanja']]

		kompbelss2C = pd.concat([kompbelss2B,kdkomp], axis=1, sort=False).sort_index()
		kompbelss2C.loc[:,'kdkomp'] = kompbelss2C.loc[:,'kdkomp'].ffill()
		kompbelss2C=kompbelss2C.dropna()
		kompbelss2C['ID']=kompbelss2C['kdkomp']+'.'+kompbelss2C['ID']
		kompbelss2D = kompbelss2C[['ID', 'Uraian', 'Belanja']]

		kompbelss3 = pd.concat([kompbelss2D,kdro], axis=1, sort=False).sort_index()
		kompbelss3.loc[:,'kdro'] = kompbelss3.loc[:,'kdro'].ffill()
		kompbelss3=kompbelss3.dropna()
		kompbelss3['ID']=kompbelss3['kdro']+'.'+kompbelss3['ID']
		kompbelss4 = kompbelss3[['ID', 'Uraian', 'Belanja']]

		kompbelss5 = pd.concat([kompbelss4,kdkro], axis=1, sort=False).sort_index()
		kompbelss5.loc[:,'kdkro'] = kompbelss5.loc[:,'kdkro'].ffill()
		kompbelss5=kompbelss5.dropna()
		kompbelss5['ID']=kompbelss5['kdkro']+'.'+kompbelss5['ID']
		kompbelss6 = kompbelss5[['ID', 'Uraian', 'Belanja']]

		kompbelss7 = pd.concat([kompbelss6,kdgiat], axis=1, sort=False).sort_index()
		kompbelss7.loc[:,'kdgiat'] = kompbelss7.loc[:,'kdgiat'].ffill()
		kompbelss7=kompbelss7.dropna()
		kompbelss7['ID']=kompbelss7['kdgiat']+'.'+kompbelss7['ID']
		kompbelss8 = kompbelss7[['ID', 'Uraian', 'Belanja']]

		kompbelss9 = pd.concat([kompbelss8,kdprog], axis=1, sort=False).sort_index()
		kompbelss9.loc[:,'kdprog'] = kompbelss9.loc[:,'kdprog'].ffill()
		kompbelss9=kompbelss9.dropna()
		kompbelss9['ID']=kompbelss9['kdprog']+'.'+kompbelss9['ID']
		kompbelss10 = kompbelss9[['ID', 'Uraian', 'Belanja']]

		# Gabung Jan-Des dengan Pagu

		rpdkomp = pd.concat([kompbel10,df3], axis=1, sort=False).sort_index()
		rpdkomp=rpdkomp.dropna()

		rpdkomp['Belanja'] = rpdkomp['Belanja'].replace(to_replace = ['51','52','53','57'], value = ['Pegawai','Barang','Modal','Bantuan Sosial'])

		rpdkompsum = rpdkomp.groupby(['ID','Uraian','Belanja']).sum()[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]
		rpdkompsum.reset_index(inplace=True)

		rpdkompsum['Pagu']= rpdkompsum[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']].sum(axis=1)
		rpdkompsum['Kode'] = rpdkompsum['ID'].str[16:19]
		rpdkompsum['Keterangan']=np.nan
		rpdkompsum['Sisa RPD']=np.nan
		rpdkompsum = rpdkompsum[['ID','Kode','Uraian', 'Belanja', 'Keterangan', 'Pagu', 'Sisa RPD','Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]


		rpdskomp = pd.concat([kompbels10,df3], axis=1, sort=False).sort_index()
		rpdskomp=rpdskomp.dropna()

		rpdskomp['Belanja'] = rpdskomp['Belanja'].replace(to_replace = ['51','52','53','57'], value = ['Pegawai','Barang','Modal','Bantuan Sosial'])

		rpdskompsum = rpdskomp.groupby(['ID','Uraian','Belanja']).sum()[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]
		rpdskompsum.reset_index(inplace=True)

		rpdskompsum['Pagu']= rpdskompsum[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']].sum(axis=1)
		rpdskompsum['Kode'] = rpdskompsum['ID'].str[20:22]
		rpdskompsum['Keterangan']=np.nan
		rpdskompsum['Sisa RPD']=np.nan
		rpdskompsum = rpdskompsum[['ID','Kode','Uraian', 'Belanja', 'Keterangan', 'Pagu', 'Sisa RPD','Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]


		rpdakun = pd.concat([kompbelss10,df3], axis=1, sort=False).sort_index()
		rpdakun=rpdakun.dropna()

		rpdakun['Belanja'] = rpdakun['Belanja'].replace(to_replace = ['51','52','53','57'], value = ['Pegawai','Barang','Modal','Bantuan Sosial'])

		rpdakunsum = rpdakun.groupby(['ID','Uraian','Belanja']).sum()[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]
		rpdakunsum.reset_index(inplace=True)

		rpdakunsum['Pagu']= rpdakunsum[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']].sum(axis=1)
		rpdakunsum['Kode'] = rpdakunsum['ID'].str[23:29]
		rpdakunsum['Keterangan']=np.nan
		rpdakunsum['Sisa RPD']=np.nan
		rpdakunsum = rpdakunsum[['ID','Kode','Uraian', 'Belanja', 'Keterangan', 'Pagu', 'Sisa RPD','Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]

		## Gabung Seluruh Data RPD per Komponen

		satker = pd.concat([urpisah,urkro,urro,rpdkompsum])
		satker.sort_values(by=['ID','Kode'],na_position='first',inplace=True,ignore_index=True)
		satker.drop(index=satker.index[0], axis=0, inplace=True)
		satker['Satker']=kode_satker

		## Ketikan RPD Terakhir per Komponen
		satker['Indeks'] = range(1, len(satker) + 1)
		satker['Indeks'] = satker['Indeks']+19 
		satker['Indeks'] = satker['Indeks'].astype(str)

		satker['Sisa RPD'].loc[~satker['Pagu'].isnull()] = '=G'+satker['Indeks']+'-SUM(I'+satker['Indeks']+':T'+satker['Indeks']+')'

		satker = satker.reindex(['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

		satker = satker.fillna('')

		## Gabung Seluruh Data RPD per Subkomponen

		satker2 = pd.concat([urpisah,urkro,urro,urkomp,rpdskompsum])
		satker2.sort_values(by=['ID','Kode'],na_position='first',inplace=True,ignore_index=True)
		satker2.drop(index=satker2.index[0], axis=0, inplace=True)
		satker2['Satker']=kode_satker

		## Ketikan RPD Terakhir per Subkomponen
		satker2['Indeks'] = range(1, len(satker2) + 1)
		satker2['Indeks'] = satker2['Indeks']+19 
		satker2['Indeks'] = satker2['Indeks'].astype(str)

		satker2['Sisa RPD'].loc[~satker2['Pagu'].isnull()] = '=G'+satker2['Indeks']+'-SUM(I'+satker2['Indeks']+':T'+satker2['Indeks']+')'

		satker2 = satker2.reindex(['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

		satker2 = satker2.fillna('')

		## Gabung Seluruh Data RPD per Akun

		satker3 = pd.concat([urpisah,urkro,urro,urkomp,urskomp,rpdakunsum])
		satker3.sort_values(by=['ID','Kode'],na_position='first',inplace=True,ignore_index=True)
		satker3.drop(index=satker3.index[0], axis=0, inplace=True)
		satker3['Satker']=kode_satker

		## Ketikan RPD Terakhir per Subkomponen
		satker3['Indeks'] = range(1, len(satker3) + 1)
		satker3['Indeks'] = satker3['Indeks']+19 
		satker3['Indeks'] = satker3['Indeks'].astype(str)

		satker3['Sisa RPD'].loc[~satker3['Pagu'].isnull()] = '=G'+satker3['Indeks']+'-SUM(I'+satker3['Indeks']+':T'+satker3['Indeks']+')'

		satker3 = satker3.reindex(['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

		satker3 = satker3.fillna('')		
		
		# Data Pagu dari DIPA Usulan
		usulan = pd.read_excel(uploaded_file, index_col=None, header=6, skipfooter=5, engine='openpyxl')
	
		dfs = usulan.iloc[:,[1, 2, 11]]

		dfs2=dfs.dropna()
		dfs2.columns = ['Kode', 'Uraian', 'pagu']

		dfs2=dfs2.loc[dfs2['Uraian'].str.len() > 1]
		dfs2[['pagu']]=dfs2[['pagu']]*1000
		dfs3=dfs2[['Kode']]
		
		# Pemisahan Kode antar level RKA KL

		idprogZ = dfs3.loc[dfs3['Kode'].str.len() == 9]
		idgiatZ = dfs3.loc[dfs3['Kode'].str.len() == 4]
		idkroZ = dfs3.loc[dfs3['Kode'].str.len() == 8]
		idroZ = dfs3.loc[dfs3['Kode'].str.len() == 7]
		idkompZ = dfs3.loc[dfs3['Kode'].str.len() == 3]
		idskompZ1 = dfs3.loc[dfs3['Kode'].str.len() == 1]
		idskompZ2 = dfs3.loc[dfs3['Kode'].str.len() == 2]
		idskompZ1['Kode']='0'+idskompZ1['Kode']
		idskompZ = pd.concat([idskompZ1,idskompZ2], axis=0)

		idakunZ = dfs2.loc[dfs2['Kode'].str.len() == 6]

		# Rename Nama Kolom id

		idprogZ.columns = ['kdprog']
		idgiatZ.columns = ['kdgiat']
		idkroZ.columns = ['kdkro']
		idroZ.columns = ['kdro']
		idkompZ.columns = ['kdkomp']
		idskompZ.columns = ['kdskomp']

		idakunZ=idakunZ.drop(['Uraian'], axis=1)
		idakunZ.columns = ['kdakun','pagu']

		# Penyesuaian Kode Program, KRO, dan RO

		kdprogZ = idprogZ['kdprog'].str[-2:]
		kdgiatZ = idgiatZ['kdgiat']
		kdkroZ = idkroZ['kdkro'].str[-3:]
		kdroZ = idroZ['kdro'].str[-3:]

		# Pembentukan RPD per Komponen
		idkompZZ = pd.concat([kdprogZ,kdgiatZ,kdkroZ,kdroZ,idkompZ], axis=1, sort=False).sort_index()

		idkompZZ.loc[:,'kdprog'] = idkompZZ.loc[:,'kdprog'].ffill()
		idkompZZ.loc[:,'kdgiat'] = idkompZZ.loc[:,'kdgiat'].ffill()
		idkompZZ.loc[:,'kdkro'] = idkompZZ.loc[:,'kdkro'].ffill()
		idkompZZ.loc[:,'kdro'] = idkompZZ.loc[:,'kdro'].ffill()
		idkompZZ=idkompZZ.dropna()

		idkompZZ['ID']=idkompZZ['kdprog']+"."+idkompZZ['kdgiat']+"."+idkompZZ['kdkro']+"."+idkompZZ['kdro']+"."+idkompZZ['kdkomp']
		kdkompZZ = idkompZZ['ID']

		sumkompZ = pd.concat([kdkompZZ,idakunZ], axis=1, sort=False).sort_index()
		sumkompZ.loc[:,'ID'] = sumkompZ.loc[:,'ID'].ffill()
		sumkompZ=sumkompZ.dropna()

		sumkompZ['ID']=sumkompZ['ID']+"."+sumkompZ['kdakun'].str[:2]
		sumkompZ=sumkompZ.drop(['kdakun'], axis=1)

		sumkompZ = sumkompZ.groupby(['ID']).sum()[['pagu']]
		sumkompZ.reset_index(inplace=True)
		sumkompZ['Pagu']= sumkompZ[['pagu']].sum(axis=1)
		sumkompZ=sumkompZ[['ID','Pagu']]

		satkerZ=satker.drop(['Pagu'], axis=1)
		satkerZZ = pd.merge(satkerZ,sumkompZ,on='ID',how='outer')
		satkerZZ = satkerZZ.reindex(['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

		# Pembentukan RPD per Subkomponen
		idskompZZ = pd.concat([kdprogZ,kdgiatZ,kdkroZ,kdroZ,idkompZ,idskompZ], axis=1, sort=False).sort_index()

		idskompZZ.loc[:,'kdprog'] = idskompZZ.loc[:,'kdprog'].ffill()
		idskompZZ.loc[:,'kdgiat'] = idskompZZ.loc[:,'kdgiat'].ffill()
		idskompZZ.loc[:,'kdkro'] = idskompZZ.loc[:,'kdkro'].ffill()
		idskompZZ.loc[:,'kdro'] = idskompZZ.loc[:,'kdro'].ffill()
		idskompZZ.loc[:,'kdkomp'] = idskompZZ.loc[:,'kdkomp'].ffill()
		idskompZZ=idskompZZ.dropna()

		idskompZZ['ID']=idskompZZ['kdprog']+"."+idskompZZ['kdgiat']+"."+idskompZZ['kdkro']+"."+idskompZZ['kdro']+"."+idskompZZ['kdkomp']+"."+idskompZZ['kdskomp']
		kdskompZZ = idskompZZ['ID']

		sumskompZ = pd.concat([kdskompZZ,idakunZ], axis=1, sort=False).sort_index()
		sumskompZ.loc[:,'ID'] = sumskompZ.loc[:,'ID'].ffill()
		sumskompZ=sumskompZ.dropna()

		sumskompZ['ID']=sumskompZ['ID']+"."+sumskompZ['kdakun'].str[:2]
		sumskompZ=sumskompZ.drop(['kdakun'], axis=1)

		sumskompZ = sumskompZ.groupby(['ID']).sum()[['pagu']]
		sumskompZ.reset_index(inplace=True)
		sumskompZ['Pagu']= sumskompZ[['pagu']].sum(axis=1)
		sumskompZ=sumskompZ[['ID','Pagu']]

		satker2Z=satker2.drop(['Pagu'], axis=1)
		satker2ZZ = pd.merge(satker2Z,sumskompZ,on='ID',how='outer')
		satker2ZZ = satker2ZZ.reindex(['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

		# Pembentukan RPD per Akun
		idakunZZ = pd.concat([kdprogZ,kdgiatZ,kdkroZ,kdroZ,idkompZ,idskompZ,idakunZ], axis=1, sort=False).sort_index()

		idakunZZ.loc[:,'kdprog'] = idakunZZ.loc[:,'kdprog'].ffill()
		idakunZZ.loc[:,'kdgiat'] = idakunZZ.loc[:,'kdgiat'].ffill()
		idakunZZ.loc[:,'kdkro'] = idakunZZ.loc[:,'kdkro'].ffill()
		idakunZZ.loc[:,'kdro'] = idakunZZ.loc[:,'kdro'].ffill()
		idakunZZ.loc[:,'kdkomp'] = idakunZZ.loc[:,'kdkomp'].ffill()
		idakunZZ.loc[:,'kdskomp'] = idakunZZ.loc[:,'kdskomp'].ffill()
		idakunZZ=idakunZZ.dropna()

		idakunZZ['ID']=idakunZZ['kdprog']+"."+idakunZZ['kdgiat']+"."+idakunZZ['kdkro']+"."+idakunZZ['kdro']+"."+idakunZZ['kdkomp']+"."+idakunZZ['kdskomp']+"."+idakunZZ['kdakun']
		kdakunZZ = idakunZZ['ID']

		sumakunZ = pd.concat([kdakunZZ,idakunZ], axis=1, sort=False).sort_index()
		sumakunZ.loc[:,'ID'] = sumakunZ.loc[:,'ID'].ffill()
		sumakunZ=sumakunZ.dropna()

		sumakunZ['ID']=sumakunZ['ID']+"."+sumakunZ['kdakun'].str[:2]
		sumakunZ=sumakunZ.drop(['kdakun'], axis=1)

		sumakunZ = sumakunZ.groupby(['ID']).sum()[['pagu']]
		sumakunZ.reset_index(inplace=True)
		sumakunZ['Pagu']= sumakunZ[['pagu']].sum(axis=1)
		sumakunZ=sumakunZ[['ID','Pagu']]

		satker3Z=satker3.drop(['Pagu'], axis=1)
		satker3ZZ = pd.merge(satker3Z,sumakunZ,on='ID',how='outer')
		satker3ZZ = satker3ZZ.reindex(['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

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

		# Mulai mengetik dafatrame ke excel sheet Komponen RPD Terakhir
		satkerZZ.to_excel(writer, sheet_name='Komponen', startrow=19, header=False, index=False)
		worksheet = writer.sheets['Komponen']

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

		worksheet.write_string('C4', '> Pengisian RPD direkomendasikan level "Komponen" saja agar lebih cepat. Karena poin pentingnya adalah total per jenis belanja tiap bulan.')
		worksheet.write_string('C5', '> Sheet "Subkomponen" dan "Akun" hanya untuk membantu Satker yang memiliki 2 jenis belanja dalam 1 Komponen atau penyusunan rencana triwulan berjalan.')
		worksheet.write_string('C6', '> Apabila kolom "Sisa RPD" berwarna merah artinya ada pergeseran pagu, harap disesuaikan RPD-nya hingga Sisa RPD tidak merah atau 0.')
		worksheet.write_string('C7', '> Revisi POK di KPA : Penyesuaian RPD antar pos harus dalam 1 bulan yang sama (1 kolom yang sama) agar Hal. III DIPA tidak berubah.')
		worksheet.write_string('C8', '> Revisi Hal. III DIPA : Penyesuaian RPD tinggal isi angka rencana bulan berjalan sd Desember.')

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

		# Mulai mengetik dafatrame ke excel sheet Subkomponen RPD Terakhir
		satker2ZZ.to_excel(writer, sheet_name='Subkomponen', startrow=19, header=False, index=False)
		worksheet2 = writer.sheets['Subkomponen']

		# Write the column headers with the defined format.
		for col_num, value in enumerate(satker.columns.values):
			worksheet2.write(9, col_num, value, header_format)

		worksheet2.set_zoom(90)
		worksheet2.set_column('A:B', None, None, {'hidden': 1})
		worksheet2.set_row(2, None, info_format)
		worksheet2.set_row(3, None, info_format)
		worksheet2.set_row(4, None, info_format)
		worksheet2.set_row(5, None, info_format)
		worksheet2.set_row(6, None, info_format)
		worksheet2.set_row(7, None, info_format)
		worksheet2.set_row(8, None, info_format)

		worksheet2.set_column_pixels('C:C', 100, kode_format)
		worksheet2.set_column_pixels('D:D', 200, uraian_format)
		worksheet2.set_column_pixels('E:F', 77, belppk_format)
		worksheet2.set_column_pixels('G:T', 110, angka_format)

		worksheet2.merge_range('A1:T1','RENCANA PENARIKAN DANA (RPD) TERAKHIR TA ____', title_format)
		worksheet2.merge_range('A2:T2','INOVASI BIDANG PPA I KANWIL DJPB BANGKA BELITUNG',subtitle_format)
		worksheet2.write_string('C3', 'Satuan Kerja')
		worksheet2.write_string('D3', ': '+nama_satker,satker_format)

		worksheet2.write_string('C4', '> Pengisian RPD direkomendasikan level "Komponen" saja agar lebih cepat. Karena poin pentingnya adalah total per jenis belanja tiap bulan.')
		worksheet2.write_string('C5', '> Sheet "Subkomponen" dan "Akun" hanya untuk membantu Satker yang memiliki 2 jenis belanja dalam 1 Komponen atau penyusunan rencana triwulan berjalan.')
		worksheet2.write_string('C6', '> Apabila kolom "Sisa RPD" berwarna merah artinya ada pergeseran pagu, harap disesuaikan RPD-nya hingga Sisa RPD tidak merah atau 0.')
		worksheet2.write_string('C7', '> Revisi POK di KPA : Penyesuaian RPD antar pos harus dalam 1 bulan yang sama (1 kolom yang sama) agar Hal. III DIPA tidak berubah.')
		worksheet2.write_string('C8', '> Revisi Hal. III DIPA : Penyesuaian RPD tinggal isi angka rencana bulan berjalan sd Desember.')

		worksheet2.merge_range('C11:T11', 'Informasi Target Penyerapan',subheader_format)
		worksheet2.merge_range('C19:T19', 'Rencana Penarikan Dana (RPD)',subheader_format)
		worksheet2.merge_range('C12:F12', 'Sisa Target Penyerapan Triwulan',sumrpd_format)
		worksheet2.merge_range('C13:F13', 'Nominal Target Penyerapan Triwulan',sumrpd_format)
		worksheet2.merge_range('C14:F14', 'Akumulasi Rencana Penarikan Dana Triwulan',sumrpd_format)
		worksheet2.merge_range('C15:F15', 'Pegawai',detilrpd_format)
		worksheet2.merge_range('C16:F16', 'Barang',detilrpd_format)
		worksheet2.merge_range('C17:F17', 'Modal',detilrpd_format)
		worksheet2.merge_range('C18:F18', 'Bantuan Sosial',detilrpd_format)
		worksheet2.merge_range('H12:H14',None,subheader_format)

		worksheet2.write_dynamic_array_formula('G12', '=G13-G14',angkasum_format)
		worksheet2.merge_range('I12:K12', None)
		worksheet2.merge_range('L12:N12', None)
		worksheet2.merge_range('O12:Q12', None)
		worksheet2.merge_range('R12:T12', None)
		worksheet2.write_dynamic_array_formula('I12', '=IF(I13-I14<0,"Sudah sesuai/melebihi target triwulan",I13-I14)',angkasum_format)
		worksheet2.write_dynamic_array_formula('L12', '=IF(L13-L14<0,"Sudah sesuai/melebihi target triwulan",L13-L14)',angkasum_format)
		worksheet2.write_dynamic_array_formula('O12', '=IF(O13-O14<0,"Sudah sesuai/melebihi target triwulan",O13-O14)',angkasum_format)
		worksheet2.write_dynamic_array_formula('R12', '=IF(R13-R14<0,"Sudah sesuai/melebihi target triwulan",R13-R14)',angkasum_format)

		worksheet2.write_dynamic_array_formula('G13', '=SUM(R13)',angkasum_format)
		worksheet2.merge_range('I13:K13', None)
		worksheet2.merge_range('L13:N13', None)
		worksheet2.merge_range('O13:Q13', None)
		worksheet2.merge_range('R13:T13', None)
		worksheet2.write_string('I13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet2.write_string('L13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet2.write_string('O13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet2.write_string('R13', 'Isi nominal target OM SPAN',angkasum_format)

		worksheet2.write_dynamic_array_formula('G14', '=SUM(R14)',angkasum_format)
		worksheet2.merge_range('I14:K14', None)
		worksheet2.merge_range('L14:N14', None)
		worksheet2.merge_range('O14:Q14', None)
		worksheet2.merge_range('R14:T14', None)
		worksheet2.write_dynamic_array_formula('I14', '=SUM(I15:K18)',angkasum_format)
		worksheet2.write_dynamic_array_formula('L14', '=SUM(I15:N18)',angkasum_format)
		worksheet2.write_dynamic_array_formula('O14', '=SUM(I15:Q18)',angkasum_format)
		worksheet2.write_dynamic_array_formula('R14', '=SUM(I15:T18)',angkasum_format)

		worksheet2.write_dynamic_array_formula('G15', '=SUMIF($E:$E,$C15,$G:$G)')
		worksheet2.write_dynamic_array_formula('H15', '=G15-SUM(I15:T15)')
		worksheet2.write_dynamic_array_formula('I15', '=SUMIF($E:$E,$C15,I:I)')
		worksheet2.write_dynamic_array_formula('J15', '=SUMIF($E:$E,$C15,J:J)')
		worksheet2.write_dynamic_array_formula('K15', '=SUMIF($E:$E,$C15,K:K)')
		worksheet2.write_dynamic_array_formula('L15', '=SUMIF($E:$E,$C15,L:L)')
		worksheet2.write_dynamic_array_formula('M15', '=SUMIF($E:$E,$C15,M:M)')
		worksheet2.write_dynamic_array_formula('N15', '=SUMIF($E:$E,$C15,N:N)')
		worksheet2.write_dynamic_array_formula('O15', '=SUMIF($E:$E,$C15,O:O)')
		worksheet2.write_dynamic_array_formula('P15', '=SUMIF($E:$E,$C15,P:P)')
		worksheet2.write_dynamic_array_formula('Q15', '=SUMIF($E:$E,$C15,Q:Q)')
		worksheet2.write_dynamic_array_formula('R15', '=SUMIF($E:$E,$C15,R:R)')
		worksheet2.write_dynamic_array_formula('S15', '=SUMIF($E:$E,$C15,S:S)')
		worksheet2.write_dynamic_array_formula('T15', '=SUMIF($E:$E,$C15,T:T)')

		worksheet2.write_dynamic_array_formula('G16', '=SUMIF($E:$E,$C16,$G:$G)')
		worksheet2.write_dynamic_array_formula('H16', '=G16-SUM(I16:T16)')
		worksheet2.write_dynamic_array_formula('I16', '=SUMIF($E:$E,$C16,I:I)')
		worksheet2.write_dynamic_array_formula('J16', '=SUMIF($E:$E,$C16,J:J)')
		worksheet2.write_dynamic_array_formula('K16', '=SUMIF($E:$E,$C16,K:K)')
		worksheet2.write_dynamic_array_formula('L16', '=SUMIF($E:$E,$C16,L:L)')
		worksheet2.write_dynamic_array_formula('M16', '=SUMIF($E:$E,$C16,M:M)')
		worksheet2.write_dynamic_array_formula('N16', '=SUMIF($E:$E,$C16,N:N)')
		worksheet2.write_dynamic_array_formula('O16', '=SUMIF($E:$E,$C16,O:O)')
		worksheet2.write_dynamic_array_formula('P16', '=SUMIF($E:$E,$C16,P:P)')
		worksheet2.write_dynamic_array_formula('Q16', '=SUMIF($E:$E,$C16,Q:Q)')
		worksheet2.write_dynamic_array_formula('R16', '=SUMIF($E:$E,$C16,R:R)')
		worksheet2.write_dynamic_array_formula('S16', '=SUMIF($E:$E,$C16,S:S)')
		worksheet2.write_dynamic_array_formula('T16', '=SUMIF($E:$E,$C16,T:T)')

		worksheet2.write_dynamic_array_formula('G17', '=SUMIF($E:$E,$C17,$G:$G)')
		worksheet2.write_dynamic_array_formula('H17', '=G17-SUM(I17:T17)')
		worksheet2.write_dynamic_array_formula('I17', '=SUMIF($E:$E,$C17,I:I)')
		worksheet2.write_dynamic_array_formula('J17', '=SUMIF($E:$E,$C17,J:J)')
		worksheet2.write_dynamic_array_formula('K17', '=SUMIF($E:$E,$C17,K:K)')
		worksheet2.write_dynamic_array_formula('L17', '=SUMIF($E:$E,$C17,L:L)')
		worksheet2.write_dynamic_array_formula('M17', '=SUMIF($E:$E,$C17,M:M)')
		worksheet2.write_dynamic_array_formula('N17', '=SUMIF($E:$E,$C17,N:N)')
		worksheet2.write_dynamic_array_formula('O17', '=SUMIF($E:$E,$C17,O:O)')
		worksheet2.write_dynamic_array_formula('P17', '=SUMIF($E:$E,$C17,P:P)')
		worksheet2.write_dynamic_array_formula('Q17', '=SUMIF($E:$E,$C17,Q:Q)')
		worksheet2.write_dynamic_array_formula('R17', '=SUMIF($E:$E,$C17,R:R)')
		worksheet2.write_dynamic_array_formula('S17', '=SUMIF($E:$E,$C17,S:S)')
		worksheet2.write_dynamic_array_formula('T17', '=SUMIF($E:$E,$C17,T:T)')

		worksheet2.write_dynamic_array_formula('G18', '=SUMIF($E:$E,$C18,$G:$G)')
		worksheet2.write_dynamic_array_formula('H18', '=G18-SUM(I18:T18)')
		worksheet2.write_dynamic_array_formula('I18', '=SUMIF($E:$E,$C18,I:I)')
		worksheet2.write_dynamic_array_formula('J18', '=SUMIF($E:$E,$C18,J:J)')
		worksheet2.write_dynamic_array_formula('K18', '=SUMIF($E:$E,$C18,K:K)')
		worksheet2.write_dynamic_array_formula('L18', '=SUMIF($E:$E,$C18,L:L)')
		worksheet2.write_dynamic_array_formula('M18', '=SUMIF($E:$E,$C18,M:M)')
		worksheet2.write_dynamic_array_formula('N18', '=SUMIF($E:$E,$C18,N:N)')
		worksheet2.write_dynamic_array_formula('O18', '=SUMIF($E:$E,$C18,O:O)')
		worksheet2.write_dynamic_array_formula('P18', '=SUMIF($E:$E,$C18,P:P)')
		worksheet2.write_dynamic_array_formula('Q18', '=SUMIF($E:$E,$C18,Q:Q)')
		worksheet2.write_dynamic_array_formula('R18', '=SUMIF($E:$E,$C18,R:R)')
		worksheet2.write_dynamic_array_formula('S18', '=SUMIF($E:$E,$C18,S:S)')
		worksheet2.write_dynamic_array_formula('T18', '=SUMIF($E:$E,$C18,T:T)')

		worksheet2.conditional_format('H15:H1048576', {'type': 'cell',
												 'criteria': '<>',
												 'value': 0,
												 'format': angkarpd_format})

		worksheet2.ignore_errors({'number_stored_as_text': 'C20:C1048576'})

		worksheet2.freeze_panes('I11')

		# Mulai mengetik dafatrame ke excel sheet Akun RPD Terakhir
		satker3ZZ.to_excel(writer, sheet_name='Akun', startrow=19, header=False, index=False)
		worksheet3 = writer.sheets['Akun']

		# Write the column headers with the defined format.
		for col_num, value in enumerate(satker.columns.values):
			worksheet3.write(9, col_num, value, header_format)

		worksheet3.set_zoom(90)
		worksheet3.set_column('A:B', None, None, {'hidden': 1})
		worksheet3.set_row(2, None, info_format)
		worksheet3.set_row(3, None, info_format)
		worksheet3.set_row(4, None, info_format)
		worksheet3.set_row(5, None, info_format)
		worksheet3.set_row(6, None, info_format)
		worksheet3.set_row(7, None, info_format)
		worksheet3.set_row(8, None, info_format)

		worksheet3.set_column_pixels('C:C', 100, kode_format)
		worksheet3.set_column_pixels('D:D', 200, uraian_format)
		worksheet3.set_column_pixels('E:F', 77, belppk_format)
		worksheet3.set_column_pixels('G:T', 110, angka_format)

		worksheet3.merge_range('A1:T1','RENCANA PENARIKAN DANA (RPD) TERAKHIR TA ____', title_format)
		worksheet3.merge_range('A2:T2','INOVASI BIDANG PPA I KANWIL DJPB BANGKA BELITUNG',subtitle_format)
		worksheet3.write_string('C3', 'Satuan Kerja')
		worksheet3.write_string('D3', ': '+nama_satker,satker_format)

		worksheet3.write_string('C4', '> Pengisian RPD direkomendasikan level "Komponen" saja agar lebih cepat. Karena poin pentingnya adalah total per jenis belanja tiap bulan.')
		worksheet3.write_string('C5', '> Sheet "Subkomponen" dan "Akun" hanya untuk membantu Satker yang memiliki 2 jenis belanja dalam 1 Komponen atau penyusunan rencana triwulan berjalan.')
		worksheet3.write_string('C6', '> Apabila kolom "Sisa RPD" berwarna merah artinya ada pergeseran pagu, harap disesuaikan RPD-nya hingga Sisa RPD tidak merah atau 0.')
		worksheet3.write_string('C7', '> Revisi POK di KPA : Penyesuaian RPD antar pos harus dalam 1 bulan yang sama (1 kolom yang sama) agar Hal. III DIPA tidak berubah.')
		worksheet3.write_string('C8', '> Revisi Hal. III DIPA : Penyesuaian RPD tinggal isi angka rencana bulan berjalan sd Desember.')


		worksheet3.merge_range('C11:T11', 'Informasi Target Penyerapan',subheader_format)
		worksheet3.merge_range('C19:T19', 'Rencana Penarikan Dana (RPD)',subheader_format)
		worksheet3.merge_range('C12:F12', 'Sisa Target Penyerapan Triwulan',sumrpd_format)
		worksheet3.merge_range('C13:F13', 'Nominal Target Penyerapan Triwulan',sumrpd_format)
		worksheet3.merge_range('C14:F14', 'Akumulasi Rencana Penarikan Dana Triwulan',sumrpd_format)
		worksheet3.merge_range('C15:F15', 'Pegawai',detilrpd_format)
		worksheet3.merge_range('C16:F16', 'Barang',detilrpd_format)
		worksheet3.merge_range('C17:F17', 'Modal',detilrpd_format)
		worksheet3.merge_range('C18:F18', 'Bantuan Sosial',detilrpd_format)
		worksheet3.merge_range('H12:H14',None,subheader_format)

		worksheet3.write_dynamic_array_formula('G12', '=G13-G14',angkasum_format)
		worksheet3.merge_range('I12:K12', None)
		worksheet3.merge_range('L12:N12', None)
		worksheet3.merge_range('O12:Q12', None)
		worksheet3.merge_range('R12:T12', None)
		worksheet3.write_dynamic_array_formula('I12', '=IF(I13-I14<0,"Sudah sesuai/melebihi target triwulan",I13-I14)',angkasum_format)
		worksheet3.write_dynamic_array_formula('L12', '=IF(L13-L14<0,"Sudah sesuai/melebihi target triwulan",L13-L14)',angkasum_format)
		worksheet3.write_dynamic_array_formula('O12', '=IF(O13-O14<0,"Sudah sesuai/melebihi target triwulan",O13-O14)',angkasum_format)
		worksheet3.write_dynamic_array_formula('R12', '=IF(R13-R14<0,"Sudah sesuai/melebihi target triwulan",R13-R14)',angkasum_format)

		worksheet3.write_dynamic_array_formula('G13', '=SUM(R13)',angkasum_format)
		worksheet3.merge_range('I13:K13', None)
		worksheet3.merge_range('L13:N13', None)
		worksheet3.merge_range('O13:Q13', None)
		worksheet3.merge_range('R13:T13', None)
		worksheet3.write_string('I13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet3.write_string('L13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet3.write_string('O13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet3.write_string('R13', 'Isi nominal target OM SPAN',angkasum_format)

		worksheet3.write_dynamic_array_formula('G14', '=SUM(R14)',angkasum_format)
		worksheet3.merge_range('I14:K14', None)
		worksheet3.merge_range('L14:N14', None)
		worksheet3.merge_range('O14:Q14', None)
		worksheet3.merge_range('R14:T14', None)
		worksheet3.write_dynamic_array_formula('I14', '=SUM(I15:K18)',angkasum_format)
		worksheet3.write_dynamic_array_formula('L14', '=SUM(I15:N18)',angkasum_format)
		worksheet3.write_dynamic_array_formula('O14', '=SUM(I15:Q18)',angkasum_format)
		worksheet3.write_dynamic_array_formula('R14', '=SUM(I15:T18)',angkasum_format)

		worksheet3.write_dynamic_array_formula('G15', '=SUMIF($E:$E,$C15,$G:$G)')
		worksheet3.write_dynamic_array_formula('H15', '=G15-SUM(I15:T15)')
		worksheet3.write_dynamic_array_formula('I15', '=SUMIF($E:$E,$C15,I:I)')
		worksheet3.write_dynamic_array_formula('J15', '=SUMIF($E:$E,$C15,J:J)')
		worksheet3.write_dynamic_array_formula('K15', '=SUMIF($E:$E,$C15,K:K)')
		worksheet3.write_dynamic_array_formula('L15', '=SUMIF($E:$E,$C15,L:L)')
		worksheet3.write_dynamic_array_formula('M15', '=SUMIF($E:$E,$C15,M:M)')
		worksheet3.write_dynamic_array_formula('N15', '=SUMIF($E:$E,$C15,N:N)')
		worksheet3.write_dynamic_array_formula('O15', '=SUMIF($E:$E,$C15,O:O)')
		worksheet3.write_dynamic_array_formula('P15', '=SUMIF($E:$E,$C15,P:P)')
		worksheet3.write_dynamic_array_formula('Q15', '=SUMIF($E:$E,$C15,Q:Q)')
		worksheet3.write_dynamic_array_formula('R15', '=SUMIF($E:$E,$C15,R:R)')
		worksheet3.write_dynamic_array_formula('S15', '=SUMIF($E:$E,$C15,S:S)')
		worksheet3.write_dynamic_array_formula('T15', '=SUMIF($E:$E,$C15,T:T)')

		worksheet3.write_dynamic_array_formula('G16', '=SUMIF($E:$E,$C16,$G:$G)')
		worksheet3.write_dynamic_array_formula('H16', '=G16-SUM(I16:T16)')
		worksheet3.write_dynamic_array_formula('I16', '=SUMIF($E:$E,$C16,I:I)')
		worksheet3.write_dynamic_array_formula('J16', '=SUMIF($E:$E,$C16,J:J)')
		worksheet3.write_dynamic_array_formula('K16', '=SUMIF($E:$E,$C16,K:K)')
		worksheet3.write_dynamic_array_formula('L16', '=SUMIF($E:$E,$C16,L:L)')
		worksheet3.write_dynamic_array_formula('M16', '=SUMIF($E:$E,$C16,M:M)')
		worksheet3.write_dynamic_array_formula('N16', '=SUMIF($E:$E,$C16,N:N)')
		worksheet3.write_dynamic_array_formula('O16', '=SUMIF($E:$E,$C16,O:O)')
		worksheet3.write_dynamic_array_formula('P16', '=SUMIF($E:$E,$C16,P:P)')
		worksheet3.write_dynamic_array_formula('Q16', '=SUMIF($E:$E,$C16,Q:Q)')
		worksheet3.write_dynamic_array_formula('R16', '=SUMIF($E:$E,$C16,R:R)')
		worksheet3.write_dynamic_array_formula('S16', '=SUMIF($E:$E,$C16,S:S)')
		worksheet3.write_dynamic_array_formula('T16', '=SUMIF($E:$E,$C16,T:T)')

		worksheet3.write_dynamic_array_formula('G17', '=SUMIF($E:$E,$C17,$G:$G)')
		worksheet3.write_dynamic_array_formula('H17', '=G17-SUM(I17:T17)')
		worksheet3.write_dynamic_array_formula('I17', '=SUMIF($E:$E,$C17,I:I)')
		worksheet3.write_dynamic_array_formula('J17', '=SUMIF($E:$E,$C17,J:J)')
		worksheet3.write_dynamic_array_formula('K17', '=SUMIF($E:$E,$C17,K:K)')
		worksheet3.write_dynamic_array_formula('L17', '=SUMIF($E:$E,$C17,L:L)')
		worksheet3.write_dynamic_array_formula('M17', '=SUMIF($E:$E,$C17,M:M)')
		worksheet3.write_dynamic_array_formula('N17', '=SUMIF($E:$E,$C17,N:N)')
		worksheet3.write_dynamic_array_formula('O17', '=SUMIF($E:$E,$C17,O:O)')
		worksheet3.write_dynamic_array_formula('P17', '=SUMIF($E:$E,$C17,P:P)')
		worksheet3.write_dynamic_array_formula('Q17', '=SUMIF($E:$E,$C17,Q:Q)')
		worksheet3.write_dynamic_array_formula('R17', '=SUMIF($E:$E,$C17,R:R)')
		worksheet3.write_dynamic_array_formula('S17', '=SUMIF($E:$E,$C17,S:S)')
		worksheet3.write_dynamic_array_formula('T17', '=SUMIF($E:$E,$C17,T:T)')

		worksheet3.write_dynamic_array_formula('G18', '=SUMIF($E:$E,$C18,$G:$G)')
		worksheet3.write_dynamic_array_formula('H18', '=G18-SUM(I18:T18)')
		worksheet3.write_dynamic_array_formula('I18', '=SUMIF($E:$E,$C18,I:I)')
		worksheet3.write_dynamic_array_formula('J18', '=SUMIF($E:$E,$C18,J:J)')
		worksheet3.write_dynamic_array_formula('K18', '=SUMIF($E:$E,$C18,K:K)')
		worksheet3.write_dynamic_array_formula('L18', '=SUMIF($E:$E,$C18,L:L)')
		worksheet3.write_dynamic_array_formula('M18', '=SUMIF($E:$E,$C18,M:M)')
		worksheet3.write_dynamic_array_formula('N18', '=SUMIF($E:$E,$C18,N:N)')
		worksheet3.write_dynamic_array_formula('O18', '=SUMIF($E:$E,$C18,O:O)')
		worksheet3.write_dynamic_array_formula('P18', '=SUMIF($E:$E,$C18,P:P)')
		worksheet3.write_dynamic_array_formula('Q18', '=SUMIF($E:$E,$C18,Q:Q)')
		worksheet3.write_dynamic_array_formula('R18', '=SUMIF($E:$E,$C18,R:R)')
		worksheet3.write_dynamic_array_formula('S18', '=SUMIF($E:$E,$C18,S:S)')
		worksheet3.write_dynamic_array_formula('T18', '=SUMIF($E:$E,$C18,T:T)')

		worksheet3.conditional_format('H15:H1048576', {'type': 'cell',
												 'criteria': '<>',
												 'value': 0,
												 'format': angkarpd_format})

		worksheet3.ignore_errors({'number_stored_as_text': 'C20:C1048576'})

		worksheet3.freeze_panes('I11')

		workbook.set_properties({
			'title':    'Perawas RPD',
			'subject':  'Perencanaan dan Pengawasan RPD',
			'author':   'Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung',
			'company':  'Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung',
			'category': 'Perencanaan Kas',
			'keywords': 'Perencanaan, Kas, Keuangan',
			'comments': 'Inovasi dari Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung'})

		writer.close()

		excelname = '[DJPb Babel] RPD Terakhir Satker ' + kode_satker + '.xlsx'
		
		st.download_button(
			label="Download RPD Terakhir Satker Anda di sini",
			data=buffer,
			file_name=excelname,
			mime="application/vnd.ms-excel"
		)

with tab2:
	st.header("Unduh RPD Realisasi untuk membantu Revisi Halaman III DIPA")
	# Upload File RPD DIPA Usulan dan Mon SAKTI
	uploaded_fileZ = st.file_uploader('Upload File RPD DIPA Usulan di sini..', type='xlsx')
	uploaded_fileZ2 = st.file_uploader('Upload File Realisasi MON SAKTI di sini.', type='xlsx')
	
	if uploaded_fileZ and uploaded_fileZ2:
		raw = pd.read_excel(uploaded_fileZ, index_col=None, header=6, skipfooter=5, engine='openpyxl')
		info = pd.read_excel(uploaded_fileZ, index_col=None, nrows = 1, dtype=str, engine='openpyxl')
		info.rename(columns = {'Unnamed: 4':'kdsatker', 'Unnamed: 6':'nmsatker'}, inplace = True)
		info['kdsatker'] = info['kdsatker'].str[-6:]
		nama_satker = info['nmsatker'].iloc[0]  +' (' + info['kdsatker'].iloc[0] + ')'
		kode_satker = info['kdsatker'].iloc[0]	

		# Master Data RPD

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
		idskomp1 = idmaster.loc[idmaster['Kode'].str.len() == 1]
		idskomp2 = idmaster.loc[idmaster['Kode'].str.len() == 2]
		idskomp1['Kode']='0'+idskomp1['Kode']
		idskomp = pd.concat([idskomp1,idskomp2], axis=0)
		idakun = idmaster.loc[idmaster['Kode'].str.len() == 6]

		# Rename Nama Kolom id

		idprog.columns = ['kdprog', 'Uraian']
		idgiat.columns = ['kdgiat', 'Uraian']
		idkro.columns = ['kdkro', 'Uraian']
		idro.columns = ['kdro', 'Uraian']
		idkomp.columns = ['kdkomp', 'Uraian']
		idskomp.columns = ['kdskomp', 'Uraian']

		# Mengambil jenis belanja saja

		idjenbel = idakun['Kode'].str[:2]
		idakun.columns = ['kdakun', 'Uraian'] #Harus rename di sini karena akan berpengaruh ke kode di bawah kalau sebelum syntax di atas

		# Penyesuaian Kode Program, KRO, dan RO

		kdprog = idprog['kdprog'].str[-2:]
		kdkro = idkro['kdkro'].str[-3:]
		kdro = idro['kdro'].str[-3:]
		kdgiat = idgiat['kdgiat']
		kdkomp = idkomp['kdkomp']
		kdskomp = idskomp['kdskomp']

		# Proses Bentuk Uraian RO & KRO

		urpisah = pd.concat([kdprog,idkro], axis=1, sort=False).sort_index()
		urpisah.loc[:,'kdprog'] = urpisah.loc[:,'kdprog'].ffill()
		urpisah=urpisah.dropna()
		urpisah['ID']=urpisah['kdprog']+'.'+urpisah['kdkro']+'.000.000.00'
		urpisah['Kode']=np.nan
		urpisah['Uraian']=np.nan
		urpisah = urpisah[['ID', 'Kode','Uraian']]
		urpisah.reset_index(drop=True, inplace=True)

		urkro = pd.concat([kdprog,idkro], axis=1, sort=False).sort_index()
		urkro.loc[:,'kdprog'] = urkro.loc[:,'kdprog'].ffill()
		urkro=urkro.dropna()
		urkro['ID']=urkro['kdprog']+'.'+urkro['kdkro']+'.000.000.00'
		urkro['Kode']=urkro['kdprog']+'.'+urkro['kdkro']
		urkro = urkro[['ID', 'Kode','Uraian']]
		urkro.reset_index(drop=True, inplace=True)

		urro = pd.concat([kdprog,kdgiat,idro], axis=1, sort=False).sort_index()
		urro.loc[:,'kdprog'] = urro.loc[:,'kdprog'].ffill()
		urro.loc[:,'kdgiat'] = urro.loc[:,'kdgiat'].ffill()
		urro=urro.dropna()
		urro['ID']=urro['kdprog']+'.'+urro['kdgiat']+'.'+urro['kdro']+'.000.00'
		urro['Kode']=urro['kdgiat']+'.'+urro['kdro']
		urro = urro[['ID', 'Kode','Uraian']]
		urro.reset_index(drop=True, inplace=True)

		urkomp = pd.concat([kdprog,kdgiat,kdkro,kdro,idkomp], axis=1, sort=False).sort_index()
		urkomp.loc[:,'kdprog'] = urkomp.loc[:,'kdprog'].ffill()
		urkomp.loc[:,'kdgiat'] = urkomp.loc[:,'kdgiat'].ffill()
		urkomp.loc[:,'kdkro'] = urkomp.loc[:,'kdkro'].ffill()
		urkomp.loc[:,'kdro'] = urkomp.loc[:,'kdro'].ffill()
		urkomp=urkomp.dropna()
		urkomp['ID']=urkomp['kdprog']+'.'+urkomp['kdgiat']+'.'+urkomp['kdkro']+'.'+urkomp['kdro']+'.'+urkomp['kdkomp']+'.00'
		urkomp['Kode']=urkomp['kdkomp']
		urkomp = urkomp[['ID', 'Kode','Uraian']]
		urkomp.reset_index(drop=True, inplace=True)

		urskomp = pd.concat([kdprog,kdgiat,kdkro,kdro,kdkomp,idskomp], axis=1, sort=False).sort_index()
		urskomp.loc[:,'kdprog'] = urskomp.loc[:,'kdprog'].ffill()
		urskomp.loc[:,'kdgiat'] = urskomp.loc[:,'kdgiat'].ffill()
		urskomp.loc[:,'kdkro'] = urskomp.loc[:,'kdkro'].ffill()
		urskomp.loc[:,'kdro'] = urskomp.loc[:,'kdro'].ffill()
		urskomp.loc[:,'kdkomp'] = urskomp.loc[:,'kdkomp'].ffill()
		urskomp=urskomp.dropna()
		urskomp['ID']=urskomp['kdprog']+'.'+urskomp['kdgiat']+'.'+urskomp['kdkro']+'.'+urskomp['kdro']+'.'+urskomp['kdkomp']+'.'+urskomp['kdskomp']+'.00'
		urskomp['Kode']=urskomp['kdskomp']
		urskomp = urskomp[['ID', 'Kode','Uraian']]
		urskomp.reset_index(drop=True, inplace=True)

		# Proses Bentuk Nilai Komponen & Jenis Belanja

		kompbel1 = pd.concat([idkomp,idjenbel], axis=1, sort=False).sort_index()

		kompbel1.loc[:,'kdkomp'] = kompbel1.loc[:,'kdkomp'].ffill()
		kompbel1.loc[:,'Uraian'] = kompbel1.loc[:,'Uraian'].ffill()
		kompbel1=kompbel1.dropna()
		kompbel1['ID']=kompbel1['kdkomp']+'.'+kompbel1['Kode']
		kompbel2 = kompbel1[['ID', 'Uraian', 'Kode']]
		kompbel2.rename(columns = {'Kode':'Belanja'}, inplace = True)

		kompbel3 = pd.concat([kompbel2,kdro], axis=1, sort=False).sort_index()
		kompbel3.loc[:,'kdro'] = kompbel3.loc[:,'kdro'].ffill()
		kompbel3=kompbel3.dropna()
		kompbel3['ID']=kompbel3['kdro']+'.'+kompbel3['ID']
		kompbel4 = kompbel3[['ID', 'Uraian', 'Belanja']]

		kompbel5 = pd.concat([kompbel4,kdkro], axis=1, sort=False).sort_index()
		kompbel5.loc[:,'kdkro'] = kompbel5.loc[:,'kdkro'].ffill()
		kompbel5=kompbel5.dropna()
		kompbel5['ID']=kompbel5['kdkro']+'.'+kompbel5['ID']
		kompbel6 = kompbel5[['ID', 'Uraian', 'Belanja']]

		kompbel7 = pd.concat([kompbel6,kdgiat], axis=1, sort=False).sort_index()
		kompbel7.loc[:,'kdgiat'] = kompbel7.loc[:,'kdgiat'].ffill()
		kompbel7=kompbel7.dropna()
		kompbel7['ID']=kompbel7['kdgiat']+'.'+kompbel7['ID']
		kompbel8 = kompbel7[['ID', 'Uraian', 'Belanja']]

		kompbel9 = pd.concat([kompbel8,kdprog], axis=1, sort=False).sort_index()
		kompbel9.loc[:,'kdprog'] = kompbel9.loc[:,'kdprog'].ffill()
		kompbel9=kompbel9.dropna()
		kompbel9['ID']=kompbel9['kdprog']+'.'+kompbel9['ID']
		kompbel10 = kompbel9[['ID', 'Uraian', 'Belanja']]

		# Proses Bentuk Nilai Subkomponen & Jenis Belanja

		kompbels1 = pd.concat([idskomp,idjenbel], axis=1, sort=False).sort_index()

		kompbels1.loc[:,'kdskomp'] = kompbels1.loc[:,'kdskomp'].ffill()
		kompbels1.loc[:,'Uraian'] = kompbels1.loc[:,'Uraian'].ffill()
		kompbels1=kompbels1.dropna()
		kompbels1['ID']=kompbels1['kdskomp']+'.'+kompbels1['Kode']
		kompbels2 = kompbels1[['ID', 'Uraian', 'Kode']]
		kompbels2.rename(columns = {'Kode':'Belanja'}, inplace = True)

		kompbels2A = pd.concat([kompbels2,kdkomp], axis=1, sort=False).sort_index()
		kompbels2A.loc[:,'kdkomp'] = kompbels2A.loc[:,'kdkomp'].ffill()
		kompbels2A=kompbels2A.dropna()
		kompbels2A['ID']=kompbels2A['kdkomp']+'.'+kompbels2A['ID']
		kompbels2B = kompbels2A[['ID', 'Uraian', 'Belanja']]

		kompbels3 = pd.concat([kompbels2B,kdro], axis=1, sort=False).sort_index()
		kompbels3.loc[:,'kdro'] = kompbels3.loc[:,'kdro'].ffill()
		kompbels3=kompbels3.dropna()
		kompbels3['ID']=kompbels3['kdro']+'.'+kompbels3['ID']
		kompbels4 = kompbels3[['ID', 'Uraian', 'Belanja']]

		kompbels5 = pd.concat([kompbels4,kdkro], axis=1, sort=False).sort_index()
		kompbels5.loc[:,'kdkro'] = kompbels5.loc[:,'kdkro'].ffill()
		kompbels5=kompbels5.dropna()
		kompbels5['ID']=kompbels5['kdkro']+'.'+kompbels5['ID']
		kompbels6 = kompbels5[['ID', 'Uraian', 'Belanja']]

		kompbels7 = pd.concat([kompbels6,kdgiat], axis=1, sort=False).sort_index()
		kompbels7.loc[:,'kdgiat'] = kompbels7.loc[:,'kdgiat'].ffill()
		kompbels7=kompbels7.dropna()
		kompbels7['ID']=kompbels7['kdgiat']+'.'+kompbels7['ID']
		kompbels8 = kompbels7[['ID', 'Uraian', 'Belanja']]

		kompbels9 = pd.concat([kompbels8,kdprog], axis=1, sort=False).sort_index()
		kompbels9.loc[:,'kdprog'] = kompbels9.loc[:,'kdprog'].ffill()
		kompbels9=kompbels9.dropna()
		kompbels9['ID']=kompbels9['kdprog']+'.'+kompbels9['ID']
		kompbels10 = kompbels9[['ID', 'Uraian', 'Belanja']]

		# Proses Bentuk Nilai Akun & Jenis Belanja

		kompbelss1 = pd.concat([idakun,idjenbel], axis=1, sort=False).sort_index()

		kompbelss1.loc[:,'kdakun'] = kompbelss1.loc[:,'kdakun'].ffill()
		kompbelss1.loc[:,'Uraian'] = kompbelss1.loc[:,'Uraian'].ffill()
		kompbelss1=kompbelss1.dropna()
		kompbelss1['ID']=kompbelss1['kdakun']+'.'+kompbelss1['Kode']
		kompbelss2 = kompbelss1[['ID', 'Uraian', 'Kode']]
		kompbelss2.rename(columns = {'Kode':'Belanja'}, inplace = True)

		kompbelss2A = pd.concat([kompbelss2,kdskomp], axis=1, sort=False).sort_index()
		kompbelss2A.loc[:,'kdskomp'] = kompbelss2A.loc[:,'kdskomp'].ffill()
		kompbelss2A=kompbelss2A.dropna()
		kompbelss2A['ID']=kompbelss2A['kdskomp']+'.'+kompbelss2A['ID']
		kompbelss2B = kompbelss2A[['ID', 'Uraian', 'Belanja']]

		kompbelss2C = pd.concat([kompbelss2B,kdkomp], axis=1, sort=False).sort_index()
		kompbelss2C.loc[:,'kdkomp'] = kompbelss2C.loc[:,'kdkomp'].ffill()
		kompbelss2C=kompbelss2C.dropna()
		kompbelss2C['ID']=kompbelss2C['kdkomp']+'.'+kompbelss2C['ID']
		kompbelss2D = kompbelss2C[['ID', 'Uraian', 'Belanja']]

		kompbelss3 = pd.concat([kompbelss2D,kdro], axis=1, sort=False).sort_index()
		kompbelss3.loc[:,'kdro'] = kompbelss3.loc[:,'kdro'].ffill()
		kompbelss3=kompbelss3.dropna()
		kompbelss3['ID']=kompbelss3['kdro']+'.'+kompbelss3['ID']
		kompbelss4 = kompbelss3[['ID', 'Uraian', 'Belanja']]

		kompbelss5 = pd.concat([kompbelss4,kdkro], axis=1, sort=False).sort_index()
		kompbelss5.loc[:,'kdkro'] = kompbelss5.loc[:,'kdkro'].ffill()
		kompbelss5=kompbelss5.dropna()
		kompbelss5['ID']=kompbelss5['kdkro']+'.'+kompbelss5['ID']
		kompbelss6 = kompbelss5[['ID', 'Uraian', 'Belanja']]

		kompbelss7 = pd.concat([kompbelss6,kdgiat], axis=1, sort=False).sort_index()
		kompbelss7.loc[:,'kdgiat'] = kompbelss7.loc[:,'kdgiat'].ffill()
		kompbelss7=kompbelss7.dropna()
		kompbelss7['ID']=kompbelss7['kdgiat']+'.'+kompbelss7['ID']
		kompbelss8 = kompbelss7[['ID', 'Uraian', 'Belanja']]

		kompbelss9 = pd.concat([kompbelss8,kdprog], axis=1, sort=False).sort_index()
		kompbelss9.loc[:,'kdprog'] = kompbelss9.loc[:,'kdprog'].ffill()
		kompbelss9=kompbelss9.dropna()
		kompbelss9['ID']=kompbelss9['kdprog']+'.'+kompbelss9['ID']
		kompbelss10 = kompbelss9[['ID', 'Uraian', 'Belanja']]

		# Gabung Jan-Des dengan Pagu

		rpdkomp = pd.concat([kompbel10,df3], axis=1, sort=False).sort_index()
		rpdkomp=rpdkomp.dropna()

		rpdkomp['Belanja'] = rpdkomp['Belanja'].replace(to_replace = ['51','52','53','57'], value = ['Pegawai','Barang','Modal','Bantuan Sosial'])

		rpdkompsum = rpdkomp.groupby(['ID','Uraian','Belanja']).sum()[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]
		rpdkompsum.reset_index(inplace=True)

		rpdkompsum['Pagu']= rpdkompsum[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']].sum(axis=1)
		rpdkompsum['Kode'] = rpdkompsum['ID'].str[16:19]
		rpdkompsum['Keterangan']=np.nan
		rpdkompsum['Sisa RPD']=np.nan
		rpdkompsum = rpdkompsum[['ID','Kode','Uraian', 'Belanja', 'Keterangan', 'Pagu', 'Sisa RPD','Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]


		rpdskomp = pd.concat([kompbels10,df3], axis=1, sort=False).sort_index()
		rpdskomp=rpdskomp.dropna()

		rpdskomp['Belanja'] = rpdskomp['Belanja'].replace(to_replace = ['51','52','53','57'], value = ['Pegawai','Barang','Modal','Bantuan Sosial'])

		rpdskompsum = rpdskomp.groupby(['ID','Uraian','Belanja']).sum()[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]
		rpdskompsum.reset_index(inplace=True)

		rpdskompsum['Pagu']= rpdskompsum[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']].sum(axis=1)
		rpdskompsum['Kode'] = rpdskompsum['ID'].str[20:22]
		rpdskompsum['Keterangan']=np.nan
		rpdskompsum['Sisa RPD']=np.nan
		rpdskompsum = rpdskompsum[['ID','Kode','Uraian', 'Belanja', 'Keterangan', 'Pagu', 'Sisa RPD','Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]


		rpdakun = pd.concat([kompbelss10,df3], axis=1, sort=False).sort_index()
		rpdakun=rpdakun.dropna()

		rpdakun['Belanja'] = rpdakun['Belanja'].replace(to_replace = ['51','52','53','57'], value = ['Pegawai','Barang','Modal','Bantuan Sosial'])

		rpdakunsum = rpdakun.groupby(['ID','Uraian','Belanja']).sum()[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]
		rpdakunsum.reset_index(inplace=True)

		rpdakunsum['Pagu']= rpdakunsum[['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']].sum(axis=1)
		rpdakunsum['Kode'] = rpdakunsum['ID'].str[23:29]
		rpdakunsum['Keterangan']=np.nan
		rpdakunsum['Sisa RPD']=np.nan
		rpdakunsum = rpdakunsum[['ID','Kode','Uraian', 'Belanja', 'Keterangan', 'Pagu', 'Sisa RPD','Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]

		## Gabung Seluruh Data RPD per Komponen

		satker = pd.concat([urpisah,urkro,urro,rpdkompsum])
		satker.sort_values(by=['ID','Kode'],na_position='first',inplace=True,ignore_index=True)
		satker.drop(index=satker.index[0], axis=0, inplace=True)
		satker['Satker']=kode_satker

		## Ketikan RPD Terakhir per Komponen
		satker['Indeks'] = range(1, len(satker) + 1)
		satker['Indeks'] = satker['Indeks']+19 
		satker['Indeks'] = satker['Indeks'].astype(str)

		satker['Sisa RPD'].loc[~satker['Pagu'].isnull()] = '=G'+satker['Indeks']+'-SUM(I'+satker['Indeks']+':T'+satker['Indeks']+')'

		satker = satker.reindex(['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

		satker = satker.fillna('')

		## Gabung Seluruh Data RPD per Subkomponen

		satker2 = pd.concat([urpisah,urkro,urro,urkomp,rpdskompsum])
		satker2.sort_values(by=['ID','Kode'],na_position='first',inplace=True,ignore_index=True)
		satker2.drop(index=satker2.index[0], axis=0, inplace=True)
		satker2['Satker']=kode_satker

		## Ketikan RPD Terakhir per Subkomponen
		satker2['Indeks'] = range(1, len(satker2) + 1)
		satker2['Indeks'] = satker2['Indeks']+19 
		satker2['Indeks'] = satker2['Indeks'].astype(str)

		satker2['Sisa RPD'].loc[~satker2['Pagu'].isnull()] = '=G'+satker2['Indeks']+'-SUM(I'+satker2['Indeks']+':T'+satker2['Indeks']+')'

		satker2 = satker2.reindex(['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

		satker2 = satker2.fillna('')

		## Gabung Seluruh Data RPD per Akun

		satker3 = pd.concat([urpisah,urkro,urro,urkomp,urskomp,rpdakunsum])
		satker3.sort_values(by=['ID','Kode'],na_position='first',inplace=True,ignore_index=True)
		satker3.drop(index=satker3.index[0], axis=0, inplace=True)
		satker3['Satker']=kode_satker

		## Ketikan RPD Terakhir per Subkomponen
		satker3['Indeks'] = range(1, len(satker3) + 1)
		satker3['Indeks'] = satker3['Indeks']+19 
		satker3['Indeks'] = satker3['Indeks'].astype(str)

		satker3['Sisa RPD'].loc[~satker3['Pagu'].isnull()] = '=G'+satker3['Indeks']+'-SUM(I'+satker3['Indeks']+':T'+satker3['Indeks']+')'

		satker3 = satker3.reindex(['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

		satker3 = satker3.fillna('')		

		# Data Realisasi Komponen
		real = pd.read_excel(uploaded_fileZ2, index_col=None, header=2, engine='openpyxl')
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
		real3 = pd.merge(idreal,real2,on='ID',how='left')
		real4 = real3.fillna(0)

		real4 = real4.reindex(['ID','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

		# Ketikan Realisasi Komponen
		rpdreal = satker[['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD',]]
		rpdreal1 = pd.merge(rpdreal,real4,on='ID',how='left')

		# Data Realisasi Subkomponen

		realZ2 = real[["TANGGAL SP2D", "KODE COA",	"NILAI RUPIAH"]]
		realZ2.rename(columns={'NILAI RUPIAH': 'nilai'}, inplace=True)
		realZ2[['satker', 'kppn', 'Akun','program','kro','sdana','bank','kewenangan','lokasi','budget','xxx','xxxx','ro','Komponen','Sub Komponen','xxxxx']] = realZ2['KODE COA'].str.split('.', expand=True)
		realZ2['ID'] = realZ2['program'].str[-2:]+"."+realZ2['kro'].str[:4]+"."+realZ2['kro'].str[-3:]+"."+realZ2['ro']+"."+realZ2['Komponen']+"."+realZ2['Sub Komponen']+"."+realZ2['Akun'].str[:2]
		realZ2['TANGGAL SP2D']=pd.to_datetime(realZ2['TANGGAL SP2D'])
		realZ2['bulan'] = realZ2['TANGGAL SP2D'].dt.month_name().str[:3]
		realZ2['Jenis Belanja'] = realZ2['Akun'].str[:2]
		realZ2 = realZ2[["ID", "bulan",	"nilai"]]
		realZ2=pd.pivot_table(realZ2, values='nilai', index='ID', columns='bulan', aggfunc='sum', fill_value=0, dropna=True, sort=True)
		realZ2.reset_index(inplace=True)
		idrealZ = rpdskompsum[['ID']]
		realZ3 = pd.merge(idrealZ,realZ2,on='ID',how='left')
		realZ4 = realZ3.fillna(0)

		realZ4 = realZ4.reindex(['ID','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

		# Ketikan Realisasi Subkomponen
		rpdrealZ = satker2[['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD',]]
		rpdreal2 = pd.merge(rpdrealZ,realZ4,on='ID',how='left')

		# Data Realisasi Akun

		realZZ2 = real[["TANGGAL SP2D", "KODE COA",	"NILAI RUPIAH"]]
		realZZ2.rename(columns={'NILAI RUPIAH': 'nilai'}, inplace=True)
		realZZ2[['satker', 'kppn', 'Akun','program','kro','sdana','bank','kewenangan','lokasi','budget','xxx','xxxx','ro','Komponen','Sub Komponen','xxxxx']] = realZZ2['KODE COA'].str.split('.', expand=True)
		realZZ2['ID'] = realZZ2['program'].str[-2:]+"."+realZZ2['kro'].str[:4]+"."+realZZ2['kro'].str[-3:]+"."+realZZ2['ro']+"."+realZZ2['Komponen']+"."+realZZ2['Sub Komponen']+"."+realZZ2['Akun']+"."+realZZ2['Akun'].str[:2]
		realZZ2['TANGGAL SP2D']=pd.to_datetime(realZZ2['TANGGAL SP2D'])
		realZZ2['bulan'] = realZZ2['TANGGAL SP2D'].dt.month_name().str[:3]
		realZZ2['Jenis Belanja'] = realZZ2['Akun'].str[:2]
		realZZ2 = realZZ2[["ID", "bulan",	"nilai"]]
		realZZ2=pd.pivot_table(realZZ2, values='nilai', index='ID', columns='bulan', aggfunc='sum', fill_value=0, dropna=True, sort=True)
		realZZ2.reset_index(inplace=True)
		idrealZZ = rpdakunsum[['ID']]
		realZZ3 = pd.merge(idrealZZ,realZZ2,on='ID',how='left')
		realZZ4 = realZZ3.fillna(0)

		realZZ4 = realZZ4.reindex(['ID','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], axis='columns')

		# Ketikan Realisasi Akun
		rpdrealZZ = satker3[['ID','Satker','Kode','Uraian','Belanja','Keterangan','Pagu','Sisa RPD',]]
		rpdreal3 = pd.merge(rpdrealZZ,realZZ4,on='ID',how='left')

		# Create a Pandas Excel writer using XlsxWriter as the engine.
		excelname = '/content/'+'[DJPb Babel] RPD Realisasi ' + kode_satker + '.xlsx'
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

		percent_format = workbook.add_format({'num_format': '0.00%'})

		# Mulai mengetik dafatrame ke excel sheet Komponen RPD Terakhir
		rpdreal1.to_excel(writer, sheet_name='Komponen', startrow=19, header=False, index=False)
		worksheet = writer.sheets['Komponen']

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

		worksheet.write_string('C4', '> Pengisian RPD direkomendasikan level "Komponen" saja agar lebih cepat. Karena poin pentingnya adalah total per jenis belanja tiap bulan.')
		worksheet.write_string('C5', '> Sheet "Subkomponen" dan "Akun" hanya untuk membantu Satker yang memiliki 2 jenis belanja dalam 1 Komponen atau penyusunan rencana triwulan berjalan.')
		worksheet.write_string('C6', '> Apabila kolom "Sisa RPD" berwarna merah artinya ada pergeseran pagu, harap disesuaikan RPD-nya hingga Sisa RPD tidak merah atau 0.')
		worksheet.write_string('C7', '> Revisi POK di KPA : Penyesuaian RPD antar pos harus dalam 1 bulan yang sama (1 kolom yang sama) agar Hal. III DIPA tidak berubah.')
		worksheet.write_string('C8', '> Revisi Hal. III DIPA : Penyesuaian RPD tinggal isi angka rencana bulan berjalan sd Desember.')

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

		# Mulai mengetik dafatrame ke excel sheet Subkomponen RPD Terakhir
		rpdreal2.to_excel(writer, sheet_name='Subkomponen', startrow=19, header=False, index=False)
		worksheet2 = writer.sheets['Subkomponen']

		# Write the column headers with the defined format.
		for col_num, value in enumerate(satker.columns.values):
			worksheet2.write(9, col_num, value, header_format)

		worksheet2.set_zoom(90)
		worksheet2.set_column('A:B', None, None, {'hidden': 1})
		worksheet2.set_row(2, None, info_format)
		worksheet2.set_row(3, None, info_format)
		worksheet2.set_row(4, None, info_format)
		worksheet2.set_row(5, None, info_format)
		worksheet2.set_row(6, None, info_format)
		worksheet2.set_row(7, None, info_format)
		worksheet2.set_row(8, None, info_format)

		worksheet2.set_column_pixels('C:C', 100, kode_format)
		worksheet2.set_column_pixels('D:D', 200, uraian_format)
		worksheet2.set_column_pixels('E:F', 77, belppk_format)
		worksheet2.set_column_pixels('G:T', 110, angka_format)

		worksheet2.merge_range('A1:T1','RENCANA PENARIKAN DANA (RPD) TERAKHIR TA ____', title_format)
		worksheet2.merge_range('A2:T2','INOVASI BIDANG PPA I KANWIL DJPB BANGKA BELITUNG',subtitle_format)
		worksheet2.write_string('C3', 'Satuan Kerja')
		worksheet2.write_string('D3', ': '+nama_satker,satker_format)

		worksheet2.write_string('C4', '> Pengisian RPD direkomendasikan level "Komponen" saja agar lebih cepat. Karena poin pentingnya adalah total per jenis belanja tiap bulan.')
		worksheet2.write_string('C5', '> Sheet "Subkomponen" dan "Akun" hanya untuk membantu Satker yang memiliki 2 jenis belanja dalam 1 Komponen atau penyusunan rencana triwulan berjalan.')
		worksheet2.write_string('C6', '> Apabila kolom "Sisa RPD" berwarna merah artinya ada pergeseran pagu, harap disesuaikan RPD-nya hingga Sisa RPD tidak merah atau 0.')
		worksheet2.write_string('C7', '> Revisi POK di KPA : Penyesuaian RPD antar pos harus dalam 1 bulan yang sama (1 kolom yang sama) agar Hal. III DIPA tidak berubah.')
		worksheet2.write_string('C8', '> Revisi Hal. III DIPA : Penyesuaian RPD tinggal isi angka rencana bulan berjalan sd Desember.')

		worksheet2.merge_range('C11:T11', 'Informasi Target Penyerapan',subheader_format)
		worksheet2.merge_range('C19:T19', 'Rencana Penarikan Dana (RPD)',subheader_format)
		worksheet2.merge_range('C12:F12', 'Sisa Target Penyerapan Triwulan',sumrpd_format)
		worksheet2.merge_range('C13:F13', 'Nominal Target Penyerapan Triwulan',sumrpd_format)
		worksheet2.merge_range('C14:F14', 'Akumulasi Rencana Penarikan Dana Triwulan',sumrpd_format)
		worksheet2.merge_range('C15:F15', 'Pegawai',detilrpd_format)
		worksheet2.merge_range('C16:F16', 'Barang',detilrpd_format)
		worksheet2.merge_range('C17:F17', 'Modal',detilrpd_format)
		worksheet2.merge_range('C18:F18', 'Bantuan Sosial',detilrpd_format)
		worksheet2.merge_range('H12:H14',None,subheader_format)

		worksheet2.write_dynamic_array_formula('G12', '=G13-G14',angkasum_format)
		worksheet2.merge_range('I12:K12', None)
		worksheet2.merge_range('L12:N12', None)
		worksheet2.merge_range('O12:Q12', None)
		worksheet2.merge_range('R12:T12', None)
		worksheet2.write_dynamic_array_formula('I12', '=IF(I13-I14<0,"Sudah sesuai/melebihi target triwulan",I13-I14)',angkasum_format)
		worksheet2.write_dynamic_array_formula('L12', '=IF(L13-L14<0,"Sudah sesuai/melebihi target triwulan",L13-L14)',angkasum_format)
		worksheet2.write_dynamic_array_formula('O12', '=IF(O13-O14<0,"Sudah sesuai/melebihi target triwulan",O13-O14)',angkasum_format)
		worksheet2.write_dynamic_array_formula('R12', '=IF(R13-R14<0,"Sudah sesuai/melebihi target triwulan",R13-R14)',angkasum_format)

		worksheet2.write_dynamic_array_formula('G13', '=SUM(R13)',angkasum_format)
		worksheet2.merge_range('I13:K13', None)
		worksheet2.merge_range('L13:N13', None)
		worksheet2.merge_range('O13:Q13', None)
		worksheet2.merge_range('R13:T13', None)
		worksheet2.write_string('I13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet2.write_string('L13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet2.write_string('O13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet2.write_string('R13', 'Isi nominal target OM SPAN',angkasum_format)

		worksheet2.write_dynamic_array_formula('G14', '=SUM(R14)',angkasum_format)
		worksheet2.merge_range('I14:K14', None)
		worksheet2.merge_range('L14:N14', None)
		worksheet2.merge_range('O14:Q14', None)
		worksheet2.merge_range('R14:T14', None)
		worksheet2.write_dynamic_array_formula('I14', '=SUM(I15:K18)',angkasum_format)
		worksheet2.write_dynamic_array_formula('L14', '=SUM(I15:N18)',angkasum_format)
		worksheet2.write_dynamic_array_formula('O14', '=SUM(I15:Q18)',angkasum_format)
		worksheet2.write_dynamic_array_formula('R14', '=SUM(I15:T18)',angkasum_format)

		worksheet2.write_dynamic_array_formula('G15', '=SUMIF($E:$E,$C15,$G:$G)')
		worksheet2.write_dynamic_array_formula('H15', '=G15-SUM(I15:T15)')
		worksheet2.write_dynamic_array_formula('I15', '=SUMIF($E:$E,$C15,I:I)')
		worksheet2.write_dynamic_array_formula('J15', '=SUMIF($E:$E,$C15,J:J)')
		worksheet2.write_dynamic_array_formula('K15', '=SUMIF($E:$E,$C15,K:K)')
		worksheet2.write_dynamic_array_formula('L15', '=SUMIF($E:$E,$C15,L:L)')
		worksheet2.write_dynamic_array_formula('M15', '=SUMIF($E:$E,$C15,M:M)')
		worksheet2.write_dynamic_array_formula('N15', '=SUMIF($E:$E,$C15,N:N)')
		worksheet2.write_dynamic_array_formula('O15', '=SUMIF($E:$E,$C15,O:O)')
		worksheet2.write_dynamic_array_formula('P15', '=SUMIF($E:$E,$C15,P:P)')
		worksheet2.write_dynamic_array_formula('Q15', '=SUMIF($E:$E,$C15,Q:Q)')
		worksheet2.write_dynamic_array_formula('R15', '=SUMIF($E:$E,$C15,R:R)')
		worksheet2.write_dynamic_array_formula('S15', '=SUMIF($E:$E,$C15,S:S)')
		worksheet2.write_dynamic_array_formula('T15', '=SUMIF($E:$E,$C15,T:T)')

		worksheet2.write_dynamic_array_formula('G16', '=SUMIF($E:$E,$C16,$G:$G)')
		worksheet2.write_dynamic_array_formula('H16', '=G16-SUM(I16:T16)')
		worksheet2.write_dynamic_array_formula('I16', '=SUMIF($E:$E,$C16,I:I)')
		worksheet2.write_dynamic_array_formula('J16', '=SUMIF($E:$E,$C16,J:J)')
		worksheet2.write_dynamic_array_formula('K16', '=SUMIF($E:$E,$C16,K:K)')
		worksheet2.write_dynamic_array_formula('L16', '=SUMIF($E:$E,$C16,L:L)')
		worksheet2.write_dynamic_array_formula('M16', '=SUMIF($E:$E,$C16,M:M)')
		worksheet2.write_dynamic_array_formula('N16', '=SUMIF($E:$E,$C16,N:N)')
		worksheet2.write_dynamic_array_formula('O16', '=SUMIF($E:$E,$C16,O:O)')
		worksheet2.write_dynamic_array_formula('P16', '=SUMIF($E:$E,$C16,P:P)')
		worksheet2.write_dynamic_array_formula('Q16', '=SUMIF($E:$E,$C16,Q:Q)')
		worksheet2.write_dynamic_array_formula('R16', '=SUMIF($E:$E,$C16,R:R)')
		worksheet2.write_dynamic_array_formula('S16', '=SUMIF($E:$E,$C16,S:S)')
		worksheet2.write_dynamic_array_formula('T16', '=SUMIF($E:$E,$C16,T:T)')

		worksheet2.write_dynamic_array_formula('G17', '=SUMIF($E:$E,$C17,$G:$G)')
		worksheet2.write_dynamic_array_formula('H17', '=G17-SUM(I17:T17)')
		worksheet2.write_dynamic_array_formula('I17', '=SUMIF($E:$E,$C17,I:I)')
		worksheet2.write_dynamic_array_formula('J17', '=SUMIF($E:$E,$C17,J:J)')
		worksheet2.write_dynamic_array_formula('K17', '=SUMIF($E:$E,$C17,K:K)')
		worksheet2.write_dynamic_array_formula('L17', '=SUMIF($E:$E,$C17,L:L)')
		worksheet2.write_dynamic_array_formula('M17', '=SUMIF($E:$E,$C17,M:M)')
		worksheet2.write_dynamic_array_formula('N17', '=SUMIF($E:$E,$C17,N:N)')
		worksheet2.write_dynamic_array_formula('O17', '=SUMIF($E:$E,$C17,O:O)')
		worksheet2.write_dynamic_array_formula('P17', '=SUMIF($E:$E,$C17,P:P)')
		worksheet2.write_dynamic_array_formula('Q17', '=SUMIF($E:$E,$C17,Q:Q)')
		worksheet2.write_dynamic_array_formula('R17', '=SUMIF($E:$E,$C17,R:R)')
		worksheet2.write_dynamic_array_formula('S17', '=SUMIF($E:$E,$C17,S:S)')
		worksheet2.write_dynamic_array_formula('T17', '=SUMIF($E:$E,$C17,T:T)')

		worksheet2.write_dynamic_array_formula('G18', '=SUMIF($E:$E,$C18,$G:$G)')
		worksheet2.write_dynamic_array_formula('H18', '=G18-SUM(I18:T18)')
		worksheet2.write_dynamic_array_formula('I18', '=SUMIF($E:$E,$C18,I:I)')
		worksheet2.write_dynamic_array_formula('J18', '=SUMIF($E:$E,$C18,J:J)')
		worksheet2.write_dynamic_array_formula('K18', '=SUMIF($E:$E,$C18,K:K)')
		worksheet2.write_dynamic_array_formula('L18', '=SUMIF($E:$E,$C18,L:L)')
		worksheet2.write_dynamic_array_formula('M18', '=SUMIF($E:$E,$C18,M:M)')
		worksheet2.write_dynamic_array_formula('N18', '=SUMIF($E:$E,$C18,N:N)')
		worksheet2.write_dynamic_array_formula('O18', '=SUMIF($E:$E,$C18,O:O)')
		worksheet2.write_dynamic_array_formula('P18', '=SUMIF($E:$E,$C18,P:P)')
		worksheet2.write_dynamic_array_formula('Q18', '=SUMIF($E:$E,$C18,Q:Q)')
		worksheet2.write_dynamic_array_formula('R18', '=SUMIF($E:$E,$C18,R:R)')
		worksheet2.write_dynamic_array_formula('S18', '=SUMIF($E:$E,$C18,S:S)')
		worksheet2.write_dynamic_array_formula('T18', '=SUMIF($E:$E,$C18,T:T)')

		worksheet2.conditional_format('H15:H1048576', {'type': 'cell',
												 'criteria': '<>',
												 'value': 0,
												 'format': angkarpd_format})

		worksheet2.ignore_errors({'number_stored_as_text': 'C20:C1048576'})

		worksheet2.freeze_panes('I11')

		# Mulai mengetik dafatrame ke excel sheet Akun RPD Terakhir
		rpdreal3.to_excel(writer, sheet_name='Akun', startrow=19, header=False, index=False)
		worksheet3 = writer.sheets['Akun']

		# Write the column headers with the defined format.
		for col_num, value in enumerate(satker.columns.values):
			worksheet3.write(9, col_num, value, header_format)

		worksheet3.set_zoom(90)
		worksheet3.set_column('A:B', None, None, {'hidden': 1})
		worksheet3.set_row(2, None, info_format)
		worksheet3.set_row(3, None, info_format)
		worksheet3.set_row(4, None, info_format)
		worksheet3.set_row(5, None, info_format)
		worksheet3.set_row(6, None, info_format)
		worksheet3.set_row(7, None, info_format)
		worksheet3.set_row(8, None, info_format)

		worksheet3.set_column_pixels('C:C', 100, kode_format)
		worksheet3.set_column_pixels('D:D', 200, uraian_format)
		worksheet3.set_column_pixels('E:F', 77, belppk_format)
		worksheet3.set_column_pixels('G:T', 110, angka_format)

		worksheet3.merge_range('A1:T1','RENCANA PENARIKAN DANA (RPD) TERAKHIR TA ____', title_format)
		worksheet3.merge_range('A2:T2','INOVASI BIDANG PPA I KANWIL DJPB BANGKA BELITUNG',subtitle_format)
		worksheet3.write_string('C3', 'Satuan Kerja')
		worksheet3.write_string('D3', ': '+nama_satker,satker_format)

		worksheet3.write_string('C4', '> Pengisian RPD direkomendasikan level "Komponen" saja agar lebih cepat. Karena poin pentingnya adalah total per jenis belanja tiap bulan.')
		worksheet3.write_string('C5', '> Sheet "Subkomponen" dan "Akun" hanya untuk membantu Satker yang memiliki 2 jenis belanja dalam 1 Komponen atau penyusunan rencana triwulan berjalan.')
		worksheet3.write_string('C6', '> Apabila kolom "Sisa RPD" berwarna merah artinya ada pergeseran pagu, harap disesuaikan RPD-nya hingga Sisa RPD tidak merah atau 0.')
		worksheet3.write_string('C7', '> Revisi POK di KPA : Penyesuaian RPD antar pos harus dalam 1 bulan yang sama (1 kolom yang sama) agar Hal. III DIPA tidak berubah.')
		worksheet3.write_string('C8', '> Revisi Hal. III DIPA : Penyesuaian RPD tinggal isi angka rencana bulan berjalan sd Desember.')


		worksheet3.merge_range('C11:T11', 'Informasi Target Penyerapan',subheader_format)
		worksheet3.merge_range('C19:T19', 'Rencana Penarikan Dana (RPD)',subheader_format)
		worksheet3.merge_range('C12:F12', 'Sisa Target Penyerapan Triwulan',sumrpd_format)
		worksheet3.merge_range('C13:F13', 'Nominal Target Penyerapan Triwulan',sumrpd_format)
		worksheet3.merge_range('C14:F14', 'Akumulasi Rencana Penarikan Dana Triwulan',sumrpd_format)
		worksheet3.merge_range('C15:F15', 'Pegawai',detilrpd_format)
		worksheet3.merge_range('C16:F16', 'Barang',detilrpd_format)
		worksheet3.merge_range('C17:F17', 'Modal',detilrpd_format)
		worksheet3.merge_range('C18:F18', 'Bantuan Sosial',detilrpd_format)
		worksheet3.merge_range('H12:H14',None,subheader_format)

		worksheet3.write_dynamic_array_formula('G12', '=G13-G14',angkasum_format)
		worksheet3.merge_range('I12:K12', None)
		worksheet3.merge_range('L12:N12', None)
		worksheet3.merge_range('O12:Q12', None)
		worksheet3.merge_range('R12:T12', None)
		worksheet3.write_dynamic_array_formula('I12', '=IF(I13-I14<0,"Sudah sesuai/melebihi target triwulan",I13-I14)',angkasum_format)
		worksheet3.write_dynamic_array_formula('L12', '=IF(L13-L14<0,"Sudah sesuai/melebihi target triwulan",L13-L14)',angkasum_format)
		worksheet3.write_dynamic_array_formula('O12', '=IF(O13-O14<0,"Sudah sesuai/melebihi target triwulan",O13-O14)',angkasum_format)
		worksheet3.write_dynamic_array_formula('R12', '=IF(R13-R14<0,"Sudah sesuai/melebihi target triwulan",R13-R14)',angkasum_format)

		worksheet3.write_dynamic_array_formula('G13', '=SUM(R13)',angkasum_format)
		worksheet3.merge_range('I13:K13', None)
		worksheet3.merge_range('L13:N13', None)
		worksheet3.merge_range('O13:Q13', None)
		worksheet3.merge_range('R13:T13', None)
		worksheet3.write_string('I13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet3.write_string('L13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet3.write_string('O13', 'Isi nominal target OM SPAN',angkasum_format)
		worksheet3.write_string('R13', 'Isi nominal target OM SPAN',angkasum_format)

		worksheet3.write_dynamic_array_formula('G14', '=SUM(R14)',angkasum_format)
		worksheet3.merge_range('I14:K14', None)
		worksheet3.merge_range('L14:N14', None)
		worksheet3.merge_range('O14:Q14', None)
		worksheet3.merge_range('R14:T14', None)
		worksheet3.write_dynamic_array_formula('I14', '=SUM(I15:K18)',angkasum_format)
		worksheet3.write_dynamic_array_formula('L14', '=SUM(I15:N18)',angkasum_format)
		worksheet3.write_dynamic_array_formula('O14', '=SUM(I15:Q18)',angkasum_format)
		worksheet3.write_dynamic_array_formula('R14', '=SUM(I15:T18)',angkasum_format)

		worksheet3.write_dynamic_array_formula('G15', '=SUMIF($E:$E,$C15,$G:$G)')
		worksheet3.write_dynamic_array_formula('H15', '=G15-SUM(I15:T15)')
		worksheet3.write_dynamic_array_formula('I15', '=SUMIF($E:$E,$C15,I:I)')
		worksheet3.write_dynamic_array_formula('J15', '=SUMIF($E:$E,$C15,J:J)')
		worksheet3.write_dynamic_array_formula('K15', '=SUMIF($E:$E,$C15,K:K)')
		worksheet3.write_dynamic_array_formula('L15', '=SUMIF($E:$E,$C15,L:L)')
		worksheet3.write_dynamic_array_formula('M15', '=SUMIF($E:$E,$C15,M:M)')
		worksheet3.write_dynamic_array_formula('N15', '=SUMIF($E:$E,$C15,N:N)')
		worksheet3.write_dynamic_array_formula('O15', '=SUMIF($E:$E,$C15,O:O)')
		worksheet3.write_dynamic_array_formula('P15', '=SUMIF($E:$E,$C15,P:P)')
		worksheet3.write_dynamic_array_formula('Q15', '=SUMIF($E:$E,$C15,Q:Q)')
		worksheet3.write_dynamic_array_formula('R15', '=SUMIF($E:$E,$C15,R:R)')
		worksheet3.write_dynamic_array_formula('S15', '=SUMIF($E:$E,$C15,S:S)')
		worksheet3.write_dynamic_array_formula('T15', '=SUMIF($E:$E,$C15,T:T)')

		worksheet3.write_dynamic_array_formula('G16', '=SUMIF($E:$E,$C16,$G:$G)')
		worksheet3.write_dynamic_array_formula('H16', '=G16-SUM(I16:T16)')
		worksheet3.write_dynamic_array_formula('I16', '=SUMIF($E:$E,$C16,I:I)')
		worksheet3.write_dynamic_array_formula('J16', '=SUMIF($E:$E,$C16,J:J)')
		worksheet3.write_dynamic_array_formula('K16', '=SUMIF($E:$E,$C16,K:K)')
		worksheet3.write_dynamic_array_formula('L16', '=SUMIF($E:$E,$C16,L:L)')
		worksheet3.write_dynamic_array_formula('M16', '=SUMIF($E:$E,$C16,M:M)')
		worksheet3.write_dynamic_array_formula('N16', '=SUMIF($E:$E,$C16,N:N)')
		worksheet3.write_dynamic_array_formula('O16', '=SUMIF($E:$E,$C16,O:O)')
		worksheet3.write_dynamic_array_formula('P16', '=SUMIF($E:$E,$C16,P:P)')
		worksheet3.write_dynamic_array_formula('Q16', '=SUMIF($E:$E,$C16,Q:Q)')
		worksheet3.write_dynamic_array_formula('R16', '=SUMIF($E:$E,$C16,R:R)')
		worksheet3.write_dynamic_array_formula('S16', '=SUMIF($E:$E,$C16,S:S)')
		worksheet3.write_dynamic_array_formula('T16', '=SUMIF($E:$E,$C16,T:T)')

		worksheet3.write_dynamic_array_formula('G17', '=SUMIF($E:$E,$C17,$G:$G)')
		worksheet3.write_dynamic_array_formula('H17', '=G17-SUM(I17:T17)')
		worksheet3.write_dynamic_array_formula('I17', '=SUMIF($E:$E,$C17,I:I)')
		worksheet3.write_dynamic_array_formula('J17', '=SUMIF($E:$E,$C17,J:J)')
		worksheet3.write_dynamic_array_formula('K17', '=SUMIF($E:$E,$C17,K:K)')
		worksheet3.write_dynamic_array_formula('L17', '=SUMIF($E:$E,$C17,L:L)')
		worksheet3.write_dynamic_array_formula('M17', '=SUMIF($E:$E,$C17,M:M)')
		worksheet3.write_dynamic_array_formula('N17', '=SUMIF($E:$E,$C17,N:N)')
		worksheet3.write_dynamic_array_formula('O17', '=SUMIF($E:$E,$C17,O:O)')
		worksheet3.write_dynamic_array_formula('P17', '=SUMIF($E:$E,$C17,P:P)')
		worksheet3.write_dynamic_array_formula('Q17', '=SUMIF($E:$E,$C17,Q:Q)')
		worksheet3.write_dynamic_array_formula('R17', '=SUMIF($E:$E,$C17,R:R)')
		worksheet3.write_dynamic_array_formula('S17', '=SUMIF($E:$E,$C17,S:S)')
		worksheet3.write_dynamic_array_formula('T17', '=SUMIF($E:$E,$C17,T:T)')

		worksheet3.write_dynamic_array_formula('G18', '=SUMIF($E:$E,$C18,$G:$G)')
		worksheet3.write_dynamic_array_formula('H18', '=G18-SUM(I18:T18)')
		worksheet3.write_dynamic_array_formula('I18', '=SUMIF($E:$E,$C18,I:I)')
		worksheet3.write_dynamic_array_formula('J18', '=SUMIF($E:$E,$C18,J:J)')
		worksheet3.write_dynamic_array_formula('K18', '=SUMIF($E:$E,$C18,K:K)')
		worksheet3.write_dynamic_array_formula('L18', '=SUMIF($E:$E,$C18,L:L)')
		worksheet3.write_dynamic_array_formula('M18', '=SUMIF($E:$E,$C18,M:M)')
		worksheet3.write_dynamic_array_formula('N18', '=SUMIF($E:$E,$C18,N:N)')
		worksheet3.write_dynamic_array_formula('O18', '=SUMIF($E:$E,$C18,O:O)')
		worksheet3.write_dynamic_array_formula('P18', '=SUMIF($E:$E,$C18,P:P)')
		worksheet3.write_dynamic_array_formula('Q18', '=SUMIF($E:$E,$C18,Q:Q)')
		worksheet3.write_dynamic_array_formula('R18', '=SUMIF($E:$E,$C18,R:R)')
		worksheet3.write_dynamic_array_formula('S18', '=SUMIF($E:$E,$C18,S:S)')
		worksheet3.write_dynamic_array_formula('T18', '=SUMIF($E:$E,$C18,T:T)')

		worksheet3.conditional_format('H15:H1048576', {'type': 'cell',
												 'criteria': '<>',
												 'value': 0,
												 'format': angkarpd_format})

		worksheet3.ignore_errors({'number_stored_as_text': 'C20:C1048576'})

		worksheet3.freeze_panes('I11')

		workbook.set_properties({
			'title':    'Perawas RPD',
			'subject':  'Perencanaan dan Pengawasan RPD',
			'author':   'Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung',
			'company':  'Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung',
			'category': 'Perencanaan Kas',
			'keywords': 'Perencanaan, Kas, Keuangan',
			'comments': 'Inovasi dari Kanwil Ditjen Perbendaharaan Provinsi Bangka Belitung'})

		writer.close()

		excelname = '[DJPb Babel] RPD Realisasi Satker ' + kode_satker + '.xlsx'
		
		st.download_button(
			label="Download RPD Realisasi Satker Anda di sini",
			data=buffer2,
			file_name=excelname,
			mime="application/vnd.ms-excel"
		)

st.markdown('---')
st.caption('Created by Farhan Ariq R. - Bidang PPA I, Kanwil DJPb Bangka Belitung 2022')
