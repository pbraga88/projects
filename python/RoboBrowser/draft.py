# -*- coding: latin1 -*-
# -*- coding: utf-8 -*-
import time
import re
from robobrowser import RoboBrowser

browser = RoboBrowser(history=True)

''' Login no endereço 192.168.0.1 '''
def connect():
    browser.open('http://192.168.0.1')
    form = browser.get_form(class_='box-content')
    form['login'].value = 'NET_0B4470'
    form['senha'].value = 'A811FC0B4470'
    browser.submit_form(form)
    time.sleep(20)
    # print(str(browser.select))
    form = browser.get_form(class_='box-content')
    form['novowifi'].value = 'teste2'
    form['novasenha'].value = 'TESTEteste'
    form['repetirsenha'] = 'TESTEteste'
    form.action = 'http://192.168.0.1/configuracao-rapida.html:-Infinity'
    browser.submit_form(form)

connect()
