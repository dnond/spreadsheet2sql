# -*- coding: utf-8 -*-


import gdata.spreadsheet.service 
import gdata.spreadsheet.text_db

import re
from datetime import datetime

class SpreadSheet2SQL:
	

	dbClient = None

	spreadsheet_name = ''

	spread_id = ''

	# デバッグ出力
	is_show_debug_message = True

	def __init__( self, email, password , spreadsheet_name,  is_show_debug_message=True ):
		self.dbClient = gdata.spreadsheet.service.SpreadsheetsService( email, password )
		self.dbClient.ProgrammaticLogin()

		self.spreadsheet_name = spreadsheet_name
		self.is_show_debug_message = is_show_debug_message

		self.spread_id = self.getSpreadSheetIdFromName( self.spreadsheet_name )

	def dp( self, message ):
		if self.is_show_debug_message:
			print message

	def getSpreadSheetIdFromName(self, spreadsheet_name):
		q = gdata.spreadsheet.service.ListQuery()
		q['title'] = self.spreadsheet_name

		feed = self.dbClient.GetSpreadsheetsFeed(query=q)
		spread_id = feed.entry[0].id.text.rsplit('/',1)[1]

		return spread_id

	

	def write2worksheet(self,  worksheet_name, text_list, sql_column_name = 'output-sql' ):
		"""
		テキストの配列を指定ワークシートに書き込む
		ワークシートの1行目1列目にsql_column_nameで指定した列名を設定しておく
		※sql_column_nameにアンダーバーは使えない
		"""
		#出力用ワークシートの取得
		q = gdata.spreadsheet.service.DocumentQuery()
		q.title = worksheet_name
		
		feed = self.dbClient.GetWorksheetsFeed( self.spread_id, query=q )
		worksheet_id = feed.entry[0].id.text.rsplit('/',1)[1]

		self.dp(  'worksheet_id:' + worksheet_id )

		#一旦削除 なぜか消し残るのでループで削除
		while True:
			feeds = self.dbClient.GetListFeed( self.spread_id, worksheet_id )
			self.dp( 'len:%s' % len( feeds.entry ))
			if len( feeds.entry ) == 0:
				break

			for each_entry in feeds.entry:
				self.dbClient.DeleteRow( each_entry )

		#書き込み
		for text in text_list:
			data = { sql_column_name: unicode( text ) }
			self.dbClient.InsertRow( data, self.spread_id, worksheet_id )



	def checkTargetWorksheet_4_MakingInsertSQL( self, worksheet_name ):
		"""
		insert文の作成に使用するワークシートかどうかチェックする
		デフォルトは英数字_-のワークシート名のみ対象とする
		"""
		if re.search(r'^[a-zA-Z0-9_\-]+$',worksheet_name) == None:
			return False
		else:
			return True
	


	def makeInsertSQL( self , is_need_addtinal_createdat_date=True, is_need_truncate=True, is_need_transaction=True):
		"""
		Insert文の作成。ワークシート毎にInsert文を作成し、配列で返す
		※ワークシート上では列名にアンダーバーを使用できないので、予め_を-に変換している想定
		"""
		feed = self.dbClient.GetWorksheetsFeed( self.spread_id )

		insert_sql_list = []
		for each_entry in feed.entry:

			# ワークシート名を取得
			worksheet_name = each_entry.content.text
			self.dp( worksheet_name )

			# 対象のワークシートかチェック　
			if not self.checkTargetWorksheet_4_MakingInsertSQL(worksheet_name):
				continue
			
			# worksheetのidを取得
			sheet_id = each_entry.id.text.rsplit('/',1)[1]

			# 列を取得
			rows = self.dbClient.GetListFeed( self.spread_id, sheet_id ).entry

			column_name_list = [] # 列名(ワークシートの1行目 アンダーバーが使えない)
			tmp_each_column_list = []
			each_data_list = [] #各セルデータのリスト
			row_sql_list = [] #列毎のSQLリスト

			for each_row in rows:
				article = gdata.spreadsheet.text_db.Record(row_entry=each_row)

				#列名のアンダーバーが消えるらしい
				if len(column_name_list) == 0:
					# 列名を取得する
					# keys_tmp = []

					# -を_に戻す
					tmp_each_column_list = [ each.replace('-', '_') for each in article.content.keys() ]

					column_name_list.extend( tmp_each_column_list )
					if is_need_addtinal_createdat_date:
						column_name_list.extend(['created_at', 'updated_at'])


				each_data_list = []
				for key, value in article.content.items():
					if value == None: 
						#空白カラムはnullとして扱う
						each_data_list.append( 'null' )
					else:
						#シングルクォーテーションで囲んでおく
						each_data_list.append( '\'%s\'' % unicode( value ) )

				# created_at、updated_at分のnowを追加する
				if is_need_addtinal_createdat_date:
					each_data_list.append( 'now()')
					each_data_list.append( 'now()')

				row_sql_list.append( "(%s)" % ','.join( each_data_list ) )

			#end of for

			# insert文の作成
			truncate = ''
			if is_need_truncate:
				truncate = 'Truncate %s ;' % worksheet_name

			insert_sql_list.append( truncate + "INSERT INTO %s (%s) VALUES %s ;" % ( 
											worksheet_name, 
											','.join( column_name_list ), 
											','.join( row_sql_list )
											)
									)
		# end of for

		if is_need_transaction:
			insert_sql_list.insert(0, 'start transaction;')
			insert_sql_list.append('commit;')

		return insert_sql_list




if __name__ == "__main__":
	import sys
	
	# 引数の取得
	argvs = sys.argv
	argc = len(argvs)

	try:
		if not argc == 4:
			print 'useage: ' + __file__ + ' email password spreadsheet_name'
			raise Exception( 'Proc Text File Not Exists!!' )

		email = argvs[1]
		passwd = argvs[2]
		spreadsheet_name = argvs[3]

		oSpreadSheet2SQL = SpreadSheet2SQL( email, passwd, spreadsheet_name )
		insert_sql_list = oSpreadSheet2SQL.makeInsertSQL()

		oSpreadSheet2SQL.write2worksheet( '出力', insert_sql_list )

	except Exception as e:
		print e.message
