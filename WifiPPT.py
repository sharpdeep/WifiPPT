# -*- coding: utf-8 -*-

__author__ = 'sharpdeep'

import webbrowser
import socket,os
import socketserver
import http.server
import time
import win32com
import pythoncom
from string import Template
from PPTControler import PPTControler

PORT = 8000
HOST = socket.gethostbyname(socket.gethostname())


class WifiPPTHandler(http.server.SimpleHTTPRequestHandler):
	def do_GET(self):
		if self.path == '/':
			with open('template/index_template.html','r',encoding='utf-8') as ft:
				message = ft.read()
				self.send_response(200)
				self.send_header("Content-type", "text/html")
				self.end_headers()
				self.wfile.write(message.encode('utf-8'))
		elif self.path == '/play':
			PPTControler().fullScreen()
			total_page = PPTControler().getActivePresentationSlideCount()
			with open('template/play_template.html','r',encoding='utf-8') as ft:
				message = (Template(ft.read()).substitute(current_page=1,total_page=total_page))
				self.send_response(200)
				self.send_header("Content-type", "text/html")
				self.end_headers()
				self.wfile.write(message.encode('utf-8'))
		elif self.path == '/nextpage':
			self.ajax(PPTControler().nextPage())
		elif self.path == '/prepage':
			self.ajax(PPTControler().prePage())
		elif self.path == '/click':
			self.ajax(PPTControler().click())
		elif '/static/image' in self.path:
			self.send_response(200)
			self.send_header('Content-type','image/png')
			self.end_headers()
			png_name = self.path.split('/')[-1]
			png_path = os.path.join('.','static','image',png_name)
			with open(png_path,'rb') as f:
				self.wfile.write(f.read())

	def ajax(self,ret_str):
		self.send_response(200)
		self.send_header('Content-type','text/plain')
		self.end_headers()
		self.wfile.write(str(ret_str).encode('utf-8'))

if __name__ == '__main__':
	with open('usage.html','w',encoding='utf-8') as f:
		with open('template/usage_template.html','r',encoding='utf-8') as ft:
			f.write(Template(ft.read()).substitute(host=HOST,port=PORT))

	httpd = socketserver.ThreadingTCPServer(('',PORT),WifiPPTHandler)
	# webbrowser.open_new_tab('usage.html')
	# webbrowser.open_new_tab('http://%s:%s'%(HOST,PORT))
	httpd.serve_forever()