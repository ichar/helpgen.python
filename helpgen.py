# -*- coding: utf-8 -*-

############################################
##   ----------------------------------   ##
##   *WEB-HELPER* SETUP FILES GENERATOR   ##
##   ----------------------------------   ##
############################################

import sys
import os
import codecs
import re
import random
import datetime
import xlrd

from types import UnicodeType, ListType, TupleType, StringType, IntType, LongType, FloatType, BooleanType
from string import strip, upper, lower

version_info = { 
    'version' : '2.46 beta', 
    'date'    : '2015-08-28', 
    'now'     : datetime.datetime.now().strftime('%Y-%m-%d'), 
    'time'    : datetime.datetime.now().strftime('%H:%M'), 
    'author'  : 'ichar' 
}
version = 'version %(version)s, %(date)s, %(author)s' % version_info
short_version = 'version %(version)s © DoorHan' % version_info
date_version = '%(now)s %(time)s' % version_info

# Журнал изменений:
# -----------------
# 2.46: добавлена функция ядра registerImageControl
# 2.45: добавлена функция ядра registerRestriction
# 2.44: добавлена функция ядра checkValidSquareRange
# 2.43: режим helperReset, секции Start, Finish
# 2.42: трассировка constSavedImage
# 2.41: добавлены переменные ядра helperFocusOn, calculatingDisabled
# 2.40: добавлена функция removeMargin
# 2.39: добавлена функция IsDateExpired
# 2.38: контроль параметров комплектации:
#       добавлена функция blockConstructByPriceType
# 2.37: контроль параметров комплектации:
#       добавлена функция blockConstructConstants
#       добавлена функция blockConstructSquare
# 2.36: контроль параметров комплектации:
#       добавлена функция blockConstructListItems
#       добавлена функция blockConstructStatus
# 2.35: добавлена переменная ядра showOnStartup
# 2.34: добавлена переменная ядра Math.PI/cos/sin
# 2.33: контроль дублирования описателей def_lines, duplicated
# 2.32: добавлена переменная ядра isItemStyle
# 2.31: ProductValue.getForm, определение: self.kind in ('CONSTANT',)
# 2.30: секция OnControl - интерактивный контроль параметров конструкции
# 2.27: добавлена функция ядра IsExist, переменная defaultPrompting
# 2.26: добавлена переменная ядра isAutoChanged
# 2.25: добавлены переменная ядра isConfirmation, confirmCode, confirmMessage
# 2.24: добавлены функции ядра IsEnabled, IsDisabled
# 2.23: опция -H (home debug options), IsWithoutDB
# 2.22: добавлены функции ядра initDefaultImage
# 2.21: добавлены функции ядра checkItemsIsValid, change2Items
# 2.20: добавлены функции ядра getObjectAsString, getObjectItemsCount, incValueInObject, addListOfMargins
# 2.19: добавлена функция ядра setStatusByKeys
#       добавлены идентификаторы: width, height, size, construct, options, codes, offset
#       добавлена функция ядра arange2Values
# 2.18: удален параметр обмена cp_wizard_id
# 2.17: добавлена функция ядра blockItems
# 2.16: добавлен параметр обмена cp_wizard_id
# 2.15: добавлена функция ядра IsPriceType
# 2.14: Logger, out: вывод символа конца строки в формате Windows (EOR)
# 2.13: режим печать исходного кода скрипта (--code): флаг IsPrintCode, функция print_code
# 2.12: добавлена функция ядра initListState
#       изменен порядок запуска секции Validation (флаг IsForcedValidation)
# 2.11: добавлены функции ядра get_compound_id, get_compound_article, blockListItems
# 2.10: добавлены функции ядра setObjectStatus, setStatusByKey, IsStringEndedWith, setValidBooleanItemByKey
# 2.09: версия модуля
# 2.08: добавлена функция ядра IsCountry
# 2.07: изменена кодировка def __repr__(self) (ProductTerm, ProductVersion)
# 2.06: публикация журнала изменений (log.txt)
# 2.05: добавлена функция ядра getListSelectedId
# 2.04: добавлена функция ядра makeDefaultRAL
# 2.03: print_locale (сравнение контента, lower/upper)
# 2.02: журнал регистрации модели 2 (model2_config)
# 2.01: make_min
# 2.00: html5-генератор
# 1.99: добавлены языки: латышский('lat'), финский('fin')
# 1.98: модель XML-обмена версии 2: добавлен параметр options
# 1.97: добавлен фиктивный тип значений DUMMY
# 1.96: добавлен тип значений параметров H-LIST-BOX (ProductValue.getForm)
# 1.95: капитализация заголовков, кроме ALERT, IMAGE, PICTURE
# 1.94: атрибут "format", стилевое оформление объекта описания: tG
# 1.93: атрибут "icon", стилевое оформление объекта описания: tT, tV
# 1.92: добавлен тип значений параметров H-LIST (ProductValue.getForm)
# 1.91: каталог изображений общего пользования common_image
# 1.90: поддержка мульти-списков (SELECT), стандартные функции: 
#       setListSelectedIndices
#       setListSelectedIndicesById
#       getListSelectedIndices
#       getListSelectedItems
# 1.89: добавлен тип значений параметров SELECT (ProductValue.getForm)
# 1.88: режимы сборки с принудительным отлючением опций отладки (779), печать в файл (mode='a+', 1754)
# 1.87: стандартные функции getValue, setValue, getSelfValue
# 1.86: добавлен тип значений параметров COMBOBOX (ProductValue.getForm)
# 1.85: добавлены функции IsFormFieldSelected, IsTrue, IsFalse, переменная ядра self
# 1.84: трассировка defaultProduct
# 1.83: в шаблоне SCRIPT_TEMPLATE изменены условия запуска секций #Validation# (!IsLoad, 338)
# 1.82: восстановление активной default-иллюстрации в секции Load (SCRIPT_TEMPLATE, 303)
# 1.81: атрибут группы (GroupID/SubGroupID) для инфотега: getForm, group_type/group_value (1473)
# 1.80: добавлена общая функция getMarkedSubstring
# 1.79: добавлена общая функция roundNumber
# 1.78: атрибут "style", стилевое оформление объекта описания: tT
# 1.77: атрибут "style", стилевое оформление объекта описания: tG, tV
# 1.76: добавлена трассировка "DEEPDEBUG: setImageURI" (410)
# 1.75: в шаблоне SCRIPT_TEMPLATE изменены условия запуска секций #Validation# (scriptExecCount > 0, 338)
# 1.74: исключена генерация скриптов для скрытых объектов (isHide, 2429)
# 1.73: разбор регулярного выражения IsSyntaxControl (rkey, 1177)
# 1.72: добавлены функции Date-Time
# 1.71: отладочная печать массива imagePath
# 1.70: смена ядра Flex(!)
# 1.67: добавлена переменная ядра defaultProduct
# 1.66: добавлена секция Term, изменен порядок следования секций: End, Prices, Trace
# 1.65: добавлены переменные для работы со слоями изображений
# 1.64: версия, date
# 1.63: файлы справочной документации (_globals['PRODUCT_INFO'])
# 1.62: опции сборки -C -U (капитализация заголовков)
# 1.61: составной комментарий (подсказка)
# 1.60: опции сборки и исполнениия (_globals['OPTIONS'])
#       PRODUCT_RELEASE
# 1.59: служебная индикация режимов исполнения (System Option): 
#       DEBUG
#       Without 1C
#       Without Restriction
# 1.58: трассировка "Вид цены номенклатуры" (priceTypeID)
# 1.57: добавлена локализация DEU (Германия)
# 1.56: локализация списковых констант
#       get_listitem - добавлен перевод значения без деления по словам
#       localize - сценарий перевода: контекст-словосочетание-фраза (функция localize_item)
# 1.55: добавлена локализация POL (Польша)
#       wizardID
# 1.54: опция OP_IS_HIDE для раздела констант
# 1.53: добавлены переменные ядра (defaultConstruct)
# 1.52: xml-тэг: <field name='Type'>IMAGE</field>
#       xml-тэг: <field name='Type'>PICTURE</field>
# 1.51: xml-тэг: <field name='Type'>ALERT</field>
#       добавлен тип значений параметров POP-UP MENU (ProductValue.getForm)
# 1.50: фильтры для полей формы
# 1.49: атрибут "номер строки описания" (line)
#       контроль многострочных комментариев (parse_script)
# 1.48: макроподстановка PRODUCT_LOCATION, локализация справочных документов
# 1.47: добавлена ActionScript функция IsStringStartedWith
# 1.46: маршрутизация каталога локализаций (country_path)
# 1.45: печать терминов для перевода на иностранные языки (--locale), только отсутствующие в словаре
# 1.44: статусы объектов описания: active, hide
# 1.43: инициализация переменных scriptExecCount(Init) и IsLoad(Load) перенесены в начало секций
# 1.42: LOCALIZED_CONSTANT_TYPES (добавлены константы с типом STRING)
# 1.41: символьные поля переведены в unicode, вывод в консоль
# 1.40: UTF-8!!! CZE
# 1.30: добавлена поддержка языковой локализации: ProductTerm, localize, _product_locale_definition
#       тип констант LIST перенесен в файл content.xml
#       общий словарь locale.xml
# 1.21: добавлена секция Main:extra
# 1.20: изменён порядок следования секций шаблона скрипта: Main и Validation

IsDebug = 0
IsDeepDebug = 0
IsShowScriptDebug = 0
IsObjectDebug = 0
IsWithoutDB = 0
IsShowOnlyGroups = 0
IsShowOnlyParameters = 0
IsShowOnlyValues = 0
IsShowOnlyInfoTags = 0
IsShowOnlyRequirements = 0
IsShowNotes = 0
IsShowLocations = 0
IsShowOnlyMargins = 0
IsShowOnlyConstants = 0
IsSyntaxControl = 0
IsDisableOutput = 0
IsPrintLocale = 0
IsPrintCode = 0
IsCapitalizeTitles = 0
IsUnCapitalizeTitles = 0

SequenceTypes = (TupleType, ListType,)
StringTypes = (UnicodeType, StringType,)
NumericTypes = (IntType, FloatType, LongType,)
DigitTypes = (IntType, FloatType, LongType, BooleanType,)

debug_object_id = None
print_to_file = None

OUTPUT_DATE_FORMAT = '%d.%m.%Y'
OUTPUT_DATE_TIME_FORMAT = '%Y-%m-%d %H:%M:%S'

SCRIPT_EXT = 'txt'
XML_EXT = 'xml'

EOL = '\n'
EOR = '\r\n'
EOI = ',\n'

LOCALE_DICT = {
    'eng' : 'rus',
}

DEFAULT_LOCALE = 'rus'
LOCALES = ('rus', 'eng', 'cze', 'chi', 'pol', 'deu', 'lat', 'fin')
LOCALE_ENG = 'eng'
LOCALE_CZE = 'cze'

product_locale = {}
locale_unique_items = []
logger = None
locale_logger = None
code_logger = None

LOCALIZED_CONSTANT_TYPES = ('MSG', 'LIST', 'STRING',)

default_encoding = 'cp1251'
default_unicode = 'utf-8'
default_path = 'configurators'
default_product_locale = 'locale.xls'

model2_config = 'm2.cfg'

PRODUCT_LOCATION = '%locale%'

# -----------------------------------
#   Шаблоны конфигурационных файлов
# -----------------------------------

FORM_TEMPLATE = \
{
    'header'  : u"""<?xml version="1.0" encoding="UTF-8"?>"""
    ,'doctype' : u"""
<!DOCTYPE document
    [
        <!ELEMENT document (system, task, description, record+)>
            <!ATTLIST   document
                        id        ID        #REQUIRED
                        type      NMTOKEN   #REQUIRED
            >
        <!ELEMENT system EMPTY>
            <!ATTLIST   system
                        id        ID        #REQUIRED
                        version   NMTOKEN   #REQUIRED
            >
        <!ELEMENT task EMPTY>
            <!ATTLIST   task
                        id        ID        #REQUIRED
                        version   NMTOKEN   #REQUIRED
            >
        <!ELEMENT description EMPTY>
            <!ATTLIST   description
                        id        ID        #REQUIRED
                        version   NMTOKEN   #REQUIRED
            >
        <!ELEMENT record (field+)>
            <!ATTLIST   record
                        id        ID        #REQUIRED
            >
        <!ELEMENT field ANY>
            <!ATTLIST   field
                        name      NMTOKEN   #REQUIRED
            >
    ]
>
"""
    ,'document_header' : u"""
<document id='variablesDescription' type='Respond'>
    <system id='DoorHan-CIS' version='1'/>
    <task id='CALCHELPER' version='2'/>
    <description id='%(id)s' version='1'/>
"""
    ,'document_footer' : u"""</document>"""
}

CONTENT_TEMPLATE = { \
    'header'  : u"""<?xml version="1.0" encoding="UTF-8"?>"""
    ,'doctype' : u"""
<!DOCTYPE document
    [
        <!ELEMENT document (system, task, description, lanquage, table+)>
            <!ATTLIST   document
                        id        ID        #REQUIRED
                        type      NMTOKEN   #REQUIRED
            >
        <!ELEMENT system EMPTY>
            <!ATTLIST   system
                        id        ID        #REQUIRED
                        version   NMTOKEN   #REQUIRED
            >
        <!ELEMENT task EMPTY>
            <!ATTLIST   task
                        id        ID        #REQUIRED
                        version   NMTOKEN   #REQUIRED
            >
        <!ELEMENT description EMPTY>
            <!ATTLIST   description
                        id        ID        #REQUIRED
                        version   NMTOKEN   #REQUIRED
            >
        <!ELEMENT lanquage EMPTY>
            <!ATTLIST   lanquage
                        id        ID        #REQUIRED
            >
        <!ELEMENT table (record+)>
            <!ATTLIST   table
                        name      NMTOKEN   #REQUIRED
            >
        <!ELEMENT record (field+)>
            <!ATTLIST   record
                        id        ID        #REQUIRED
            >
        <!ELEMENT field ANY>
            <!ATTLIST   field
                        name      NMTOKEN   #REQUIRED
            >
    ]
>
"""
    ,'document_header' : u"""
<document id='variablesContent' type='Respond'>
    <system id='DoorHan-CIS' version='1'/>
    <task id='CALCHELPER' version='2'/>
    <description id='%(id)s' version='1'/>
    <lanquage id='%(lang)s'/>
"""
    ,'document_footer' : u"""</document>"""
}

JAVASCRIPT_TEMPLATE = {
    'Constant' : u"""
var %(id)s = %(value)s;""",

    'Field' : u"""
var %(id)s = %(value)s;""",

    'Download' : u"""
%(indent)s%(id)s = $Field.get('%(id)s');""",

    'Upload' : u"""
%(indent)s$Field.set('%(id)s', %(id)s);""",

    'Body' : u"""%(header)s
// Global IDs & Attrs
// ------------------
var gids = new Array(%(gids)s);
var gattrs = {
%(gattrs)s
};

// Constants
// ---------
%(vars)s

// Fields
// ------
%(fields)s

// Init
// ----
$(document).ready(function() {
    $Init();
    //alert('OK');
});

// Main
// ----
function MainBrains() {
%(brains)s
}

%(footer)s""",
}

SCRIPT_TEMPLATE = { \
    'Start' : u"""
var IsDebug:Boolean = new Boolean(%(IsDebug)s);
var IsDeepDebug:Boolean = new Boolean(%(IsDeepDebug)s);
var IsLoad:Boolean = new Boolean(false);
var IsForcedValidation:Boolean = new Boolean(false);

var validation_disable:Boolean = new Boolean(%(validation_disable)s);
var prices_disable:Boolean = new Boolean(%(prices_disable)s);

if (!(IsExist('helperReset') && helperReset)) 
{""",

    'Globals' : u"""
// ------------------------------------
//  #S.Globals: ГЛОБАЛЬНЫЕ ОПРЕДЕЛЕНИЯ
// ------------------------------------

var country_path:String = '';
var now:Date = new Date();

isWorkWithout1C = %(isWorkWithout1C)s;

var common_image:String = '';
var image:String = 'default.jpg';

NoticeMessage = "";

%(code)s

//
//  ========================================================================
//

if (IsDebug) NoticeMessage += "DEBUG: scriptExecCount [" + scriptExecCount + "]<br>";
if (IsDebug) NoticeMessage += "DEBUG: changedFormFieldID [" + changedFormFieldID + "]<br>";
if (IsDebug) NoticeMessage += "DEBUG: isConfirmation [" + isConfirmation + "]<br>";
if (IsDebug) NoticeMessage += "DEBUG: confirmCode [" + confirmCode + "]<br>";
if (IsDebug) NoticeMessage += "DEBUG: withoutRestriction [" + withoutRestriction + "]<br>";
if (IsDebug) NoticeMessage += "DEBUG: currentLocale [" + currentLocale + "]<br>";
if (IsDebug) NoticeMessage += "DEBUG: defaultProduct [" + defaultProduct + "]<br>";
if (IsDebug) NoticeMessage += "DEBUG: priceTypeID [" + priceTypeID + "]<br>";
if (IsDebug) NoticeMessage += "DEBUG: helperVersion [" + helperVersion + "]<br>";
if (IsDebug) NoticeMessage += "DEBUG: wizardID [" + wizardID + "]<br>";"""

    ,'Init' : u"""
// -------------------------------------------------
//  #S.Init: ИНИЦИАЛИЗАЦИЯ (ПАРАМЕТРЫ ПО УМОЛЧАНИЮ)
// -------------------------------------------------

if (!scriptExecCount)
{
    scriptExecCount = 0;

%(code_then)s%(code)s
}
else"""

    ,'Load' : u"""
// ---------------------------------
//  #S.Load: ЗАГРУЗКА/РАСЧЕТ ЗАКАЗА
// ---------------------------------

if (scriptExecCount > 0 && !changedFormFieldID)
{
    IsLoad = true;

%(code_then)s%(code)s

    if (constSavedImage && constSavedImage != defaultImage && scriptExecCount > 0) image = constSavedImage;
}
else
{
%(code_else)s
}"""

    ,'Begin' : u"""
// ---------------------------------
//  #S.Begin: НАЧАЛЬНЫЕ ОПРЕДЕЛЕНИЯ
// ---------------------------------

%(code)s"""

    ,'OnChange' : u"""
// --------------------------------------
//  #S.OnChange: ОБРАБОТКА СОБЫТИЙ ФОРМЫ
// --------------------------------------

if (changedFormFieldID)
{
%(code_then)s%(code)s
}
else
{
%(code_else)s
}
%(code_extra)s"""

    ,'OnControl' : u"""
// -------------------------------------------------
//  #S.OnControl: ИНТЕРАКТИВНЫЙ КОНТРОЛЬ ПАРАМЕТРОВ
// -------------------------------------------------

if (scriptExecCount > 0 && IsExist('helperProductControl') && helperProductControl['active'])
{
%(code_then)s%(code)s
}"""

    ,'Validation' : u"""
// ------------------------------------
//  #S.Validation: КОНТРОЛЬ ПАРАМЕТРОВ
// ------------------------------------

if ((IsForcedValidation || (scriptExecCount > 0 && !IsLoad)) && !validation_disable && !withoutRestriction)
{
%(code_then)s%(code)s
}"""

    ,'Main' : u"""
// --------------------------------
//  #S.Main: УСТАНОВКА КОНСТРУКЦИИ
// --------------------------------

if (helperErrorCode == 0)
{
%(code_then)s%(code)s
}

// ------------------------------------
//  #S.Extra: ДОПОЛНИТЕЛЬНАЯ ОБРАБОТКА
// ------------------------------------

%(code_extra)s"""

    ,'Prices' : u"""
// ----------------------------
//  #S.Prices: ЦЕНООБРАЗОВАНИЕ
// ----------------------------

if (!prices_disable && helperErrorCode == 0)
{
%(code_then)s%(code)s
}"""

    ,'End' : u"""
// --------------------
//  #S.End: ЗАВЕРШЕНИЕ
// --------------------

%(code)s"""

    ,'Trace' : u"""
// -----------------------
//  #S.Trace: ТРАССИРОВКА
// -----------------------

var IsTrace:Boolean = new Boolean(%(on)s);

if (IsTrace) {
    NoticeMessage = %(messages)s
        "***<br>" +
    NoticeMessage;
}
%(code)s
"""

    ,'Term' : u"""
//
//  =================
//

%(code)s

if (IsDeepDebug) {
    NoticeMessage += 
        "DEEPDEBUG: total images:" + imagePath.length + "<br>" +
        "DEEPDEBUG: imagePath:" + "<br>";
    for (var i=0; i<imagePath.length; i++) {
        NoticeMessage += "==> [" + imagePath[i] + "]<br>";
    }
    NoticeMessage += 
        "DEEPDEBUG: imageVisibility:" + imageVisibility.toString() + "<br>" +
        "DEEPDEBUG: imageX:" + imageX.toString() + "<br>" +
        "DEEPDEBUG: imageY:" + imageY.toString() + "<br>";
}

setImageURI = common_image ? 'images/' + common_image + '/' + image : country_path + default_image + image;

if (IsDeepDebug) {
    NoticeMessage += 
        "DEEPDEBUG: setImageURI:" + setImageURI + "<br>";
}

if (IsDeepDebug) NoticeMessage += "DEEPDEBUG: constSavedImage [" + constSavedImage + "]<br>";"""

    ,'Finish' : u"""
}

if (prices_disable) attentionMessage = getSystemOption('prices_disable', currentLocale) + attentionMessage;
if (validation_disable) attentionMessage = getSystemOption('validation_disable', currentLocale) + attentionMessage;
if (isWorkWithout1C) attentionMessage = getSystemOption('isWorkWithout1C', currentLocale) + attentionMessage;
if (withoutRestriction) attentionMessage = getSystemOption('withoutRestriction', currentLocale) + attentionMessage;
if (IsDebug || IsDeepDebug) attentionMessage = getSystemOption('Debug', currentLocale) + attentionMessage;

++scriptExecCount;"""
}

# -----------------------------
#   Опции генератора контента
# -----------------------------

OP_RECORD_DELETE = ('D', 'd', '-',) # запись удалена
OP_IS_TRACE = ('T', 't',)           # режим автотрассировки
OP_IS_HIDE = ('H', 'h',)            # STATUS=HIDE
OP_WEB = ('+',)                     # публикация

# ---------------------------------------------------
#   Блок дескриптора "Конструкция изделия" (лист 1)
# ---------------------------------------------------

##  типы объектов описания
tG = 'G' # группа
tP = 'P' # параметр
tV = 'V' # значение
tT = 'T' # инфотег
tR = 'R' # требование (!)
tN = 'N' # подсказка

##  указатель заголовков дескриптора
ID_LINE = {'sheet': 1, 'row': 1, 'id_col': 9, 'options_col': 11, 'info_col': 17}

##  структура полей блока
S1_COLUMNS_LINE = (
    None, 
    'num', 'group', 'parameter', 'value', 'article', 'price', 
    None, None, 
    'id', 'id_1c', 'kind', 'type', 'format', 'current_value', 'unit_id', 
    None,
    'script',
    None, 
    'illustration', 'style', 'icon',
)

##  допустимые разновидности элементов описания
VALID_KIND_VALUES = (
    'CHECKBOX',
    'COMBOBOX',
    'CONSTANT',
    'DISPLAY FIELD',
    'DUMMY',
    'H-LIST-BOX',
    'H-LIST',
    'INPUT AREA',
    'INPUT FIELD',
    'LABEL',
    'POPUP',
    'PROGRESS',
    'POP-UP MENU',
    'RADIOBUTTON',
    'SELECT',
    'SIMPLE INPUT',
    'SINGLE LIST',
)

##  допустимые типы данных
VALID_TYPE_VALUES = (
    'ALERT',
    'BOOLEAN',
    'LIST',
    'IMAGE',
    'NUMBER',
    'STRING',
    'PICTURE',
)

# -----------------------------------------------
#   Блок дескриптора "Ценообразование" (лист 2)
# -----------------------------------------------

##  типы объектов описания
tM = 'M' # наценка

##  структура полей блока
S2_COLUMNS_LINE = (
    None, 
    'num', 'parameter', 'article', 'price', 
    None, None, 
    'script', 
)

# -----------------------------------------
#   Блок дескриптора "Константы" (лист 3)
# -----------------------------------------

##  типы объектов описания
tC = 'C' # константа

##  структура полей блока
S3_COLUMNS_LINE = (
    None, 
    'num', 'type', 'kind', 'id', 'value1', 'value2', 'description',
)

# ------------------------------------------------
#   Блок дескриптора "Словарь терминов" (лист 5)
# ------------------------------------------------

##  типы объектов описания
tC = 'L' # словарный термин

##  структура полей блока
S5_COLUMNS_LINE = (
    None, 
    'num', 'rus', 'eng', 'cze', 'chi', 'pol', 'deu', 'lat', 'fin'
)

# ----------------------------------------------
#   Блок дескриптора "Лист изменений" (лист 6)
# ----------------------------------------------

##  структура полей блока
S6_COLUMNS_LINE = (
    None, 
    'date', 'version', 'description', 'files',
)

##  допустимые виды констант
VALID_CONSTANT_KINDS = (
    'CONSTANT',
    'PROGRESS',
)

##  допустимые типы полей констант
VALID_CONSTANT_TYPES = (
    'BOOLEAN',
    'LIST',
    'NUMBER',
    'MSG',
    'STRING',
)

_globals = {
    'PRODUCT_ID': '', 
    'PRODUCT_TITLE': '', 
    'PRODUCT_INFO': [], 
    'PRODUCT_PATH': '',
    'PRODUCT_IMAGES_PATH': '',

    # список и последовательность вывода атрибутов XML-тегов
    'TAG_IDS_SORTED': ( \
        'ID', 'Kind', 'Type', 'Format', 'GroupID', 'SubGroupID', 'CurrentValue', 'UnitID', 'TypeID', 'Label', 'Description', 'Illustration', 'Style', 'Icon',
        'image', 'description', 'location',
    ),

    # список и последовательность вывода разделов скрипта обработчика
    'SCRIPT_SECTIONS_SORTED': ( \
        'Start', 'Globals', 'Init', 'Load', 'Begin', 'OnChange', 'OnControl', 'Main', 'Validation', 'End', 'Prices', 'Trace', 'Term', 'Finish',
    ),

    # разделители списка
    'LIST_INDEX_SEPARATOR' : '##',
    'LIST_ITEM_SEPARATOR' : '|',
    'LIST_VALUE_SEPARATOR' : ';',
}

product_lexeme = set()
as_lang = [ \
    ## языковые конструкции (ключевые слова ActionScript3, in lower case only!)
    'array',
    'boolean',
    'break',
    'case',
    'ceil',
    'cell',
    'continue',
    'cos',
    'date',
    'else',
    'false',
    'for',
    'floor',
    'function',
    'if',
    'in',
    'indexof',
    'int',
    'isnan',
    'join',
    'length',
    'max',
    'min',
    'new',
    'null',
    'number',
    'parsefloat',
    'parseint',
    'pi',
    'pop',
    'push',
    'return',
    'replace',
    'sin',
    'sort',
    'split',
    'string',
    'substr',
    'substring',
    'switch',
    'tostring',
    'touppercase',
    'tolowercase',
    'true',
    'uint',
    'valueof',
    'var',
    'while',
    ]

product_declarations = set([ \
    ## предопределенные идентификаторы шаблона
    'common_image',
    'image',
    'IsDebug',
    'IsDeepDebug',
    'IsEval',
    'IsForcedValidation',
    'IsLoad',
    'prices_disable',
    'validation_disable',

    ## общие соглашения (аргументы функций, вспомогательные переменные ...)
    'article',
    'code',
    'codes',
    'construct',
    'count',
    'default',
    'destination',
    'enabled',
    'filter',
    'finalize',
    'force',
    'height',
    'id',
    'index',
    'item',
    'item_default',
    'item_panel',
    'item_window',
    'items',
    'key',
    'lst',
    'name',
    'now',
    'ob',
    'offset',
    'options',
    'outer',
    'prices',
    'selected_id',
    'selected_index',
    'status',
    'value',
    'msg',
    's',
    'size',
    'source',
    'width',
    'with_color',
    'x',
    'y',
    'z',

    ## локализация
    'word_shield_body',
    'word_more_less',

    ## переменные Flex-ядра
    'attentionMessage',
    'calculatingDisabled',
    'changedFormFieldID',
    'confirmCode',
    'confirmMessage',
    'convertObjectToString',
    'country_path',
    'countryID',
    'currencyName',
    'currentLocale',
    'defaultConstruct',
    'defaultConstructCount',
    'defaultProduct',
    'defaultPrompting',
    'helperErrorCode',
    'helperErrorMessage',
    'helperFocusOn',
    'helperHostName',
    'helperLoadedVersion',
    'helperPageLocation',
    'helperRequestURI',
    'helperReset',
    'helperPathURI',
    'helperVersion',
    'helperSessionID',
    'imagePath',
    'imageVisibility',
    'imageX',
    'imageY',
    'isAutoChanged',
    'isConfirmation',
    'isItemStyle',
    'isWorkWithout1C',
    'NoticeMessage',
    'objectStatus',
    'options',
    'priceTypeID',
    'regionID',
    'scriptExecCount',
    'self',
    'setImageURI',
    'showOnStartup',
    'userID',
    'userTypeID',
    'withoutRestriction',
    'wizardID',
    'wizardName',

    ## общие функции (functions.txt)
    'addListItem',
    'addListOfMargins',
    'addMargin',
    'addExtraMargin',
    'applyFilter',
    'arange2Values',
    'blockConstructByPriceType',
    'blockConstructConstants',
    'blockConstructListItems',
    'blockConstructSquare',
    'blockConstructStatus',
    'blockItems',
    'blockListItem',
    'blockListItems',
    'change2Items',
    'checkItemsIsValid',
    'checkValidSquareRange',
    'cloneListItem',
    'get_compound_article',
    'get_compound_id',
    'getDate', 
    'getDay', 
    'getFilter',
    'getFilterByIntersection',
    'getFormFieldValue',
    'getFullYear', 
    'getList',
    'getListCurrIndex',
    'getListCurrValue',
    'getListItem',
    'getListItemById',
    'getListLength',
    'getListMapping',
    'getListSelectedId',
    'getListSelectedIndex',
    'getListSelectedIndices',
    'getListSelectedItem',
    'getListSelectedItems',
    'getListSelectedValue',
    'getListValidId',
    'getListValidName',
    'getMarkedSubstring',
    'getMinutes', 
    'getMonth', 
    'getObjectAsString',
    'getObjectItemsCount',
    'getSeconds',
    'getTime', 
    'getSelfValue',
    'getValue',
    'incValueInObject',
    'initDefaultImage',
    'initListState',
    'iterList',
    'IsCountry',
    'IsDateExpired',
    'IsDisabled',
    'IsEnabled',
    'IsExist',
    'IsFalse',
    'IsFormFieldSelected',
    'IsPriceType',
    'IsStringEndedWith',
    'IsStringStartedWith',
    'IsTrue',
    'makeDecimal',
    'makeDefaultRAL',
    'makeOuterList',
    'makeRAL',
    'myRound',
    'Object',
    'setListSelectedId',
    'setListSelectedIndex',
    'setListSelectedIndexById',
    'setListSelectedIndices',
    'setListSelectedIndicesById',
    'setObjectStatus',
    'setStatus',
    'setStatusByKey',
    'setStatusByKeys',
    'setValidBooleanItem',
    'setValidBooleanItemByKey',
    'setValue',
    'strip',
    'registerImageControl',
    'registerRestriction',
    'removeListItem',
    'removeListItemById',
    'removeMargin',
    'roundDecimal',
    'roundInteger',
    'roundNumber',
    'setNoticeMessage',

    ## предопределенные константы
    'LST_INDEX_DELIMETER',
    'LST_INDICES_DELIMETER',
    'LST_ITEM_DELIMETER',
    'LST_VALUE_DELIMETER',
    ])

STATUS_ACTIVE = 'A'
STATUS_HIDE = 'H'

STATUSES = {'active' : STATUS_ACTIVE, 'hide' : STATUS_HIDE}
ALERT_CONTENT = '<![CDATA[%s]]>'

def_lines = {}
duplicated = {'total' : 0, 'ids' : [], 'items' : {}}

# -------------------------
#   Информация о продукте
# -------------------------

def product_info(name, index=1):
    book = xlrd.open_workbook(name)
    sheet = book.sheet_by_index(index)

    # ID изделия
    x = sheet.cell_value(ID_LINE['row'], ID_LINE['id_col']) or None
    _globals['PRODUCT_ID'] = x

    # Опции
    x = sheet.cell_value(ID_LINE['row'], ID_LINE['options_col']).split(':') or []
    e = IsDebug or IsDeepDebug
    _globals['OPTIONS'] = {
            'IsDebug' : e and len(x) > 0 and x[0] and int(x[0]) and 'true' or 'false',
            'IsDeepDebug' : e and len(x) > 1 and x[1] and int(x[1]) and 'true' or 'false',
            'isWorkWithout1C' : (IsWithoutDB or (len(x) > 2 and x[2] and int(x[2]))) and 'true' or 'false',
            'validation_disable' : len(x) > 3 and x[3] and int(x[3]) and 'true' or 'false',
            'prices_disable' : len(x) > 4 and x[4] and int(x[4]) and 'true' or 'false',
        }

    # Заголовки описания
    x = sheet.cell_value(ID_LINE['row'], ID_LINE['info_col']).strip()
    if type(x) in StringTypes:
        x = x.encode(default_encoding, 'ignore')

    r = re.compile(r'#([\w]+)#\{(.*?)\}', re.I+re.DOTALL)
    info = []

    while x:
        m = r.search(x)
        if m is not None:
            if m.groups()[0] == 'Title':
                _globals['PRODUCT_TITLE'] = m.groups()[1]
            elif m.groups()[0] == 'Info':
                image, description, location = m.groups()[1].split(',')
                if image and description and location:
                    info.append((image.strip(), description.strip(), location.strip(),))
            elif m.groups()[0] == 'Path':
                _globals['PRODUCT_PATH'] = m.groups()[1]
            x = x[m.end():]
        else:
            break
    _globals['PRODUCT_INFO'] = info

    del sheet
    del book

def win(s):
    # Implements the dictionary
    for key, value in LOCALE_DICT.items():
        s = re.sub(r'(\[.*)%s(.*\])' % key, r'\1%s\2' % value, s)
    # Set valid line-feeds
    s = re.sub(r'[ \t\r]*\n', r'\r\n', s)
    s = re.sub(r'\r\n\r\n\r\n', r'\r\n\r\n', s)
    return s

def Capitalize(s):
    return (s and len(s) > 1 and s[0].upper() + s[1:]) or (len(s) == 1 and s.upper()) or ''

def unCapitalize(s):
    return (s and len(s) > 1 and s[0].lower() + s[1:]) or (len(s) == 1 and s.lower()) or ''

def print_locale(x):
    x = x.strip()
    if x.endswith('.'):
        x = x[:-1]
    if not (x in locale_unique_items or x in product_locale) and x != x.upper():
        locale_logger.out(x)
        locale_unique_items.append(x)

def print_code(line, x, type=None):
    #if type and type in (tM,):
    #    return
    code_logger.out('----- %s' % line, EOR)
    code_logger.out(re.sub(EOL, EOR, x), EOR)

def localize(key, locale):
    #
    # перевод словосочетания с прописной первой буквой
    #
    def localize_item(v):
        if v:
            IsUpper = v[0].isupper()
            v = unCapitalize(v)
            if v in product_locale:
                v = product_locale[v].get_value(locale)
                #if locale == 'deu':
                #    logger.out('==> %s' % v)
            else:
                if IsPrintLocale:
                    print_locale(x)
            if IsUpper:
                v = Capitalize(v)
        return v
    #
    # разбор и перевод контекста
    #
    if type(key) in StringTypes:
        if not key:
            return key or ''
        if locale == DEFAULT_LOCALE:
            return key
        if key.endswith('.'):
            IsPoint = True
            key = key[:-1]
        else:
            IsPoint = False
        value = ''
        if unCapitalize(key) in product_locale:
            # перевод словосочетания целиком
            value = localize_item(key)
        else:
            # либо по фразам (строкам), разделенным запятой
            for line in key.splitlines():
                if value:
                    value += '\n'
                if not line:
                    continue
                s = ''
                for x in line.strip().split(', '):
                    v = localize_item(x.strip())
                    if s:
                        s += ', '
                    s += v
                value += s
        # точка в конце контекста
        if IsPoint:
            value += '.'
        return value or key or ''
    else:
        return key

def make_xmltag(id, indent='', inline=False, encoding=default_unicode, **kw):
    """
        Returns XML-tag content.

        Arguments:

            id -- record tag id
            
            indent -- indent for tag line
            
            inline -- Boolean, output field content inline
            
            encoding -- body encoding, default: <default_encoding> value
            
            **kw -- attributes dict (fields): {key : value, ...}
    """
    body = value = u""
    if encoding is None:
        encoding = default_unicode
    for key in _globals['TAG_IDS_SORTED']:
        try:
            if not kw.has_key(key):
                continue
            if kw[key]:
                x = type(kw[key])
                if x in DigitTypes:
                    value = kw[key]
                elif x is UnicodeType:
                    value = kw[key]
                elif x is StringType:
                    try:
                        value = unicode(kw[key], encoding)
                    except:
                        value = kw[key].decode(default_encoding)
            else:
                value = u""
            if encoding:
                key = unicode(key, encoding)
            if inline:
                body += u"<field name='%s'>%s</field>" % (key, value)
            else:
                body += u"%s%s<field name='%s'>%s</field>" % (EOL, indent+' '*4, key, value)
        except:
            #print '--> value:', value.decode(default_encoding, 'ignore'), type(value)
            #print '--> body:', body.decode(default_encoding, 'ignore'), type(body)
            #print '--> key:', key.decode(default_encoding, 'ignore'), type(key)
            #print '--> kw[key]:', kw[key].decode(default_encoding, 'replace'), type(kw[key])
            raise
        
    return \
        inline and \
            u"%s<record id='%s'>%s%s%s%s%s</record>" % (indent, id, EOL, indent+' '*4, body, EOL, indent) \
        or \
            u"%s<record id='%s'>%s%s%s</record>" % (indent, id, body, EOL, indent)

def replaceQuotedValues(value, obs):
    msg = value
    r = re.compile(r'%#([\w]+)#%')
    m = r.search(msg)
    while m:
        if m.group(1):
            id = m.group(1)
            x = id in obs and obs[id].value or ''
            msg = msg[:m.start()] + x + msg[m.end():]
        m = r.search(msg, m.start())
    return msg

def make_min(name='script', encoding=default_unicode):
    script_file = '%s.min.js' % name
    so = Logger(script_file, encoding, bom=False)
    so.out('// %s.js' % _globals['version']['name'])
    so.out('// %s' % _globals['version']['description'])
    so.out('// %s' % _globals['version']['date'])

    source = file('%s.js' % name, 'rb')
    b = source.read()
    source.close()

    b = re.sub(r'%s*' % codecs.BOM_UTF8, '', b)
    b = re.sub(r'\/\/(.*)\n', '', b)
    b = re.sub(r'(continue|break)\s+(\w)', r'\1;\2', b)
    b = re.sub(r'[\r]+', '', b)
    b = re.sub(r'\t', ' ', b)
    b = re.sub(r'\s+({)\s+', r'\1', b)
    b = re.sub(r'\s+(})\s+', r'\1', b)
    b = re.sub(r'if\s+(\()', r'if\1', b)
    #b = re.sub(r'\;\s+(if|var|else|return|function)', r';\1', b)
    b = re.sub(r'\;\s+(\w)', r';\1', b)
    b = re.sub(r'\,\s+(\'|\")', r',\1', b)
    b = re.sub(r'\s+(=|\?|:|>=|<=|>|<|\+|\-|\:|\*|\|\||&&|==|!=)\s+', r'\1', b)
    b = re.sub(r'[\n]+', '', b)
    b = re.sub(r'\s+', ' ', b)
    #b = re.sub(r'(\;|\})(function)', r'\1 \2', b)

    so.out(b)

    so.close()
    del so

class ProductAbstract(object):
    """
        Абстрактный класс описания
    """
    metatype = None
    default_xml_indent = ' '*8
    
    def __init__(self, id, title):
        self.num = None                 # порядковый номер
        self.title = title or ''        # заголовок
        self.comment = None             # комментарий
        self.id = id                    # ID
        self.id_1c = None               # ID для связи с 1C
        self.kind = None                # разновидность элемента управления
        self.type = None                # тип данных
        self.format = None              # формат поля
        self.current_value = None       # значение по умолчанию
        self.unit_id = None             # единицы измерения
        self.script = {}                # скрипт обработчика
        self.illustration = None        # иллюстрация (вспомогательное изображение)
        self.style = None               # стиль оформления
        self.icon = None                # пиктограмма

        self._attrs = {}                # атрибуты описаний (артикул, код наценки и др)

        self.description = ''           # комментарий к требованию
        self.note = None                # подсказка

        self.objects = {}               # словарь экземпляров описания
        self.ids = []                   # структура описания
        self.parents = []               # список родителей экземпляра описания

        self.parent = None              # прямой предок

        self.is_trace = None            # режим автотрассировки

        product_declarations.add(id)    # зарегистрировать идентификатор в списке валидных лексем

        self.status = STATUS_ACTIVE
        self.line = None

    def set_line(self, line):
        if not line:
            return
        elif type(line) in NumericTypes:
            self.line = line + 1
        else:
            self.line = line

    def set_status(self, status):
        if status in STATUSES:
            self.status = STATUSES[status]

    def get_status(self):
        if self.parents:
            ob = self
            while ob:
                ob = ob.parent or None
                if ob and ob.status == STATUS_HIDE:
                    return ob.status
        return self.status

    def _init_state(self, data):
        if not data:
            return

        # зарегистрировать описатель
        if self.id not in def_lines:
            def_lines[self.id] = self.line

        if data.has_key('num') and data['num']:
            self.num = data['num']
        if data.has_key('article') and data['article']:
            self._attrs['article'] = unicode(data['article'])
        if data.has_key('price') and data['price']:
            self._attrs['price'] = unicode(data['price'])
        if data.has_key('id_1c') and data['id_1c']:
            self.id_1c = data['id_1c']
        if data.has_key('kind') and data['kind']:
            if upper(data['kind']) in VALID_KIND_VALUES:
                self.kind = data['kind']
        if data.has_key('type') and data['type']:
            if upper(data['type']) in VALID_TYPE_VALUES:
                self.type = data['type']
        if data.has_key('format') and data['format']:
            self.format = data['format']
        if data.has_key('current_value'):
            x = data['current_value']
            if unicode(x).strip():
                if self.type in ('NUMBER',):
                    if type(x) is FloatType and x-int(x) == 0:
                        # целое число
                        self.current_value = int(x)
                    else:
                        # дробное число
                        self.current_value = x * 1.0
                else:
                    self.current_value = x
            else:
                self.current_value = None
        if data.has_key('unit_id') and data['unit_id']:
            self.unit_id = data['unit_id']
        if data.has_key('script') and data['script']:
            if IsPrintCode:
                print_code(self.line, data['script'], self.get_type())
            for section, key, value in self.parse_script(data['script']):
                #
                # Структура контейнера: {<section> : {<key> : <script>, ...}, ...}, 
                # допустимые секции: 'code', 'else', 'then', 'extra' и др.
                # секция по умолчанию: 'code'
                #
                if not section in self.script:
                    self.script[section] = {}
                if not key in self.script[section]:
                    self.script[section][key] = ''
                elif self.script[section][key]:
                    self.script[section][key] += EOL
                self.script[section][key] += value
        if data.has_key('illustration') and data['illustration']:
            self.illustration = data['illustration']
        if data.has_key('style') and data['style']:
            self.style = data['style']
        if data.has_key('icon') and data['icon']:
            self.icon = data['icon']

    def __repr__(self):
        return self.id or ''

    def isActive(self):
        return self.get_status() == STATUS_ACTIVE and True or False

    def isHide(self):
        return self.get_status() == STATUS_HIDE and True or False

    def isValid(self):
        return self.id and True or False

    def isTrace(self):
        return IsDebug and self.id and self.is_trace and True or False

    def isXMLOutput(self):
        return True

    def isOverridable(self):
        return True

    def group(self):
        return ''

    def subgroup(self):
        return ''
    ##
    ##  Output interface ----------------------------
    ##
    def getForm(self, indent=None, encoding=None, locale=None):
        pass

    def getContent(self, indent=None, encoding=None, locale=None):
        pass

    def getScript(self, section=None):
        if not (section and self.script and section in self.script):
            return None
        res = []
        for key in self.script[section]:
            if key in ('then', 'else',):
                res.append((key, EOL.join(['%s%s' % (' '*4, line) for line in self.script[section][key].split(EOL)]),))
            else:
                res.append((key, self.script[section][key],))
        return res
    ##
    ##  Attribute's parsing and evolution -----------
    ##
    def set(self, ob):
        if ob is not None:
            id = ob.id
            line = ob.line
            if id:
                if id not in self.ids:
                    self.objects[id] = ob
                    self.ids.append(id)
                if id in def_lines and def_lines[id] != line:
                    if id not in duplicated['ids']:
                        duplicated['ids'].append(id)
                        duplicated['items'][id] = [def_lines[id],]
                    if line not in duplicated['items'][id]:
                        duplicated['items'][id].append(line)
                    duplicated['total'] += 1

    def get(self, id):
        assert id in self.ids, "Object is not present: [%s]" % id
        return self.objects[id]

    def get_html_title(self, locale=None):
        return localize(self.title, locale) or ''

    def get_html_value(self, locale=None):
        # Locale depended types
        if self.type in ('ALERT', 'IMAGE', 'LIST'):
            return "''"
        # and other
        if self.type == 'BOOLEAN':
            return '%s' % (self.current_value == 'true' and 'true' or 'false')
        elif self.type == 'NUMBER':
            return '%s' % str(self.current_value or 0)
        elif self.type in ('STRING', 'MSG',):
            return "''"
        else:
            return "null"

    def set_parent(self, obs):
        for ob in obs:
            if ob is None or not repr(ob):
                continue
            self.parents.append(ob.id)
            if self.parent is None:
                self.parent = ob

    def get_type(self):
        return self.metatype

    def set_description(self, value):
        self.description = value

    def set_note(self, value):
        self.note = re.sub(r'.*?\:\s+(.*)', r'\1', re.sub(r'(<.*>)', '', value)) or None

    def parse_title(self, value):
        #
        # Структура заголовка: <title{<note>}>[<comment>]
        # Может содержать заголовок, подсказку и комментарий.
        #
        r = re.compile(r'(.*)\s+\[(.*)\]')
        if not isinstance(value, basestring):
            value = unicode(value)
        try:
            m = r.search(value)
        except:
            print value
            raise
        #m = r.search(value)
        if m is not None:
            title, comment = m.groups()
        else:
            r = re.compile(r'\[(.*)\]')
            m = r.search(value)
            if m is not None:
                title, comment = None, m.groups()[0]
            else:
                title, comment = value, None
        note = ''
        if title and '{' in title and '}' in title:
            #
            # Составной комментарий (подсказка). 
            # Может состоять из нескольких фраз, заключенных в фигурные скобки
            #
            r = re.compile(r'\{(.*?)\}')
            m = None
            while True:
                m = r.search(title, m and m.end() or 0)
                if m is None:
                    break
                note += '%s%s' % (note and ' ' or '', m.groups()[0])
            if note and note.strip()[-1] not in '.,;!?':
                note += '.'
            #
            # Заголовок, из которого текст подсказки исключается
            #
            title = re.sub(r'(\{.*?\})(?si)', ' ', title)
            title = re.sub(r'([\s]+)', ' ', title).strip()
        #
        # Капитализация заголовков
        #
        if title and self.get_type() in (tV, tT,) and self.type not in ('ALERT', 'IMAGE', 'PICTURE',):
            if IsCapitalizeTitles:
                title = Capitalize(title)
            elif IsUnCapitalizeTitles:
                title = unCapitalize(title)
        return (title, note, comment,)

    def parse_script(self, value):
        #
        # Структура скрипта: #<key>:<condition>#[<{>|<""">]<body>[<}>|<""">]
        # Результат: [(<key>, <condition>, <body>), ...], где <condition>: then|else|code|extra
        #
        r1 = re.compile(r'#([\w\:]+)#"""\n*(.*?)\n*"""', re.I+re.DOTALL)
        r2 = re.compile(r'#([\w\:]+)#\{(.*?)\}', re.I+re.DOTALL)
        x = value[:]
        res = []

        while x:
            n = 0
            for r in (r1, r2,):
                m = r.search(x)
                if m is not None:
                    if m.groups()[1]:
                        keys = m.groups()[0].split(':')
                        if len(keys) == 1:
                            keys.append('code')
                        res.append((keys[0], keys[1], m.groups()[1]))
                    x = x[m.end():]
                    n += 1
            if not n:
                break
        #
        # Макроподстановка: %article% - артикул, %price% - код наценки
        #
        res = map(lambda (key, condition, script): ( \
                    key, \
                    condition, \
                    re.sub(r'\%article\%', self._attrs.get('article', ''), re.sub(r'\%price\%', self._attrs.get('price', ''), script)) \
                 ), res)
        #
        # Синтаксический контроль кода скрипта
        #
        if IsSyntaxControl:
            rkey = re.compile(r'([^\s\=\(\)\[\]])*?([\w]+?)([\s\;\<\>\=\[\]\&\|\!\(\)\,$])')
            rfunc = re.compile(r'function\s+([\w]+)\s*\(')
            rvar = re.compile(r'var\s+([\w]+)[\:\s\=\;]')
            rcom_multiline = re.compile(r'/\*(.*?)\*/(?s)')
            rcom_singleline = re.compile(r'//(.*)\n')
            rstring_const = re.compile(r'[\'\"](.*?)[\'\"]')

            message_not_valid_comment = '(!) Error [not valid multiline comment inside the script]: %s, line=%s'

            for key, condition, script in res:
                if not script:
                    continue

                # контроль многострочных комментариев
                m = None
                while True:
                    m = rcom_multiline.search(script, m and m.end() or 0)
                    if not (m and m.group()):
                        break
                    x = m.groups()[0]
                    if not x:
                        break
                    if '/*' in x or '*/' in x:
                        logger.out(message_not_valid_comment % (self.id, self.line))

                # удалить строковые константы и комментарии
                s = re.sub(rstring_const, '', re.sub(rcom_singleline, '', re.sub(rcom_multiline, '', script)))

                # финальный контроль многострочных комментариев
                if '/*' in s or '*/' in s:
                    logger.out(message_not_valid_comment % (self.id, self.line))

                # отбор используемых идентификаторов
                m = None
                while True:
                    m = rkey.search(s, m and m.end() or 0)
                    if not (m and m.group()):
                        break
                    x = m.groups()[1]
                    if not x:
                        break
                    if lower(x) not in as_lang and not x.isdigit():
                        product_lexeme.add(x)

                # поиск объявлений функций
                m = None
                while True:
                    m = rfunc.search(s, m and m.end() or 0)
                    if not (m and m.group()):
                        break
                    x = m.groups()[0]
                    if x:
                        product_declarations.add(x)

                # поиск объявлений переменных
                m = None
                while True:
                    m = rvar.search(s, m and m.end() or 0)
                    if not (m and m.group()):
                        break
                    x = m.groups()[0]
                    if x:
                        product_declarations.add(x)

        return res

class ProductGroup(ProductAbstract):
    """
        Раздел параметров изделия
    """
    metatype = tG
    
    def __init__(self, id, title=None):
        super(ProductGroup, self).__init__(id, title)

    def _init_state(self, data):
        super(ProductGroup, self)._init_state(data)
        if data['group'] and not data['num']:
            self.title, x, self.comment = self.parse_title(data['group'])

    def isValid(self):
        return self.title and self.id and True or False

    def isXMLOutput(self):
        return self.title and True or False

    def getForm(self, indent=None, encoding=None, locale=None):
        kw = {}
        if self.isXMLOutput() and self.style:
            kw.update({ \
                'Kind'         : 'GROUP',
                'Style'        : self.style,
            })
        else:
            return None
        return make_xmltag(self.id, indent=indent or self.default_xml_indent, encoding=encoding, **kw)

    def getContent(self, indent=None, encoding=None, locale=None):
        kw = {}
        if self.isXMLOutput():
            kw.update({ \
                'TypeID'       : 1, 
                'Description'  : localize(self.title, locale),
            })
        else:
            return None
        if self.format:
            kw.update({ \
                'Format'       : self.format
            })
        return make_xmltag(self.id, indent=indent or self.default_xml_indent, encoding=encoding, **kw)

class ProductParameter(ProductAbstract):
    """
        Параметр изделия
    """
    metatype = tP
    
    def __init__(self, id, title=None):
        super(ProductParameter, self).__init__(id, title)

    def _init_state(self, data):
        super(ProductParameter, self)._init_state(data)
        if data['parameter'] and data['num']:
            self.title, x, self.comment = self.parse_title(data['parameter'])

    def isValid(self):
        return self.num and self.id and True or False

    def isXMLOutput(self):
        return self.title and True or False

    def getForm(self, indent=None, encoding=None, locale=None):
        pass

    def getContent(self, indent=None, encoding=None, locale=None):
        kw = {}
        if self.isXMLOutput():
            kw.update({ \
                'TypeID'       : 2, 
                'Description'  : localize(self.title, locale),
            })
        else:
            return None
        return make_xmltag(self.id, indent=indent or self.default_xml_indent, encoding=encoding, **kw)

class ProductValue(ProductAbstract):
    """
        Значение параметра изделия
    """
    metatype = tV
    
    def __init__(self, id, title=None):
        super(ProductValue, self).__init__(id, title)

    def _init_state(self, data):
        super(ProductValue, self)._init_state(data)
        if data['value'] and not data['num'] and not data['parameter']:
            self.title, self.note, self.comment = self.parse_title(data['value'])

    def isValid(self):
        return (self.title or \
                    self.parent is not None and self.parent.title or \
                    self.parent.parent is not None and self.parent.parent.title) and \
                self.id and True or False

    def group(self):
        return self.parent.parent is not None and self.parent.parent.id or ''

    def subgroup(self):
        return self.parent is not None and self.parent.id or ''

    def get_html_title(self, locale=None):
        if self.kind == 'CHECKBOX':
            title = self.title or self.parent.title or self.parent.objects[self.parent.ids[0]].title or ''
            return '%s|%s' % (localize(title, locale), localize(u'Да', locale))
        if self.kind == 'RADIOBUTTON':
            title = self.parent.title or \
                (self.parent.ids[0] != self.id and self.parent.objects[self.parent.ids[0]].kind != 'RADIOBUTTON' and \
                    self.parent.objects[self.parent.ids[0]].title) or \
                self.parent.parent.title or ''
            return '%s|%s' % (localize(title, locale), localize(self.title, locale))
        return self.title and localize(self.title, locale) or super(ProductValue, self).get_html_title(locale)

    def getForm(self, indent=None, encoding=None, locale=None):
        if not self.isXMLOutput():
            return None
        kw = {}
        if self.kind in ('RADIOBUTTON', 'CHECKBOX', 'INPUT FIELD', 'INPUT AREA', 'POP-UP MENU', 'COMBOBOX', 'SELECT', 'H-LIST', 'H-LIST-BOX', 'DUMMY',):
            kw.update({ \
                'Kind'         : self.kind,
                'Type'         : self.type,
                'GroupID'      : self.group(),
                'SubGroupID'   : self.subgroup(),
            })
        elif self.kind in ('SIMPLE INPUT',):
            kw.update({ \
                'Kind'         : self.isOverridable() and 'INPUT FIELD' or self.kind,
                'Type'         : self.type,
                'GroupID'      : self.group(),
            })
        elif self.kind in ('DISPLAY FIELD',):
            kw.update({ \
                'Kind'         : self.kind,
                'Type'         : self.type,
                'GroupID'      : self.group(),
            })
        elif self.kind in ('LABEL',):
            kw.update({ \
                'Kind'         : self.isOverridable() and 'DISPLAY FIELD' or self.kind,
                'Type'         : self.type,
                'GroupID'      : self.group() or self.parent is not None and self.parent.get_type() == tG and self.parent.id or '',
            })
        elif self.kind in ('CONSTANT',):
            kw.update({ \
                'Kind'         : self.kind,
                'Type'         : self.type,
            })
        if self.id_1c:
            kw.update({ \
                'ID'           : self.id_1c 
            })
        if self.format:
            kw.update({ \
                'Format'       : self.format
            })
        if self.current_value is not None:
            kw.update({ \
                'CurrentValue' : self.current_value or '',
            })
        if self.unit_id:
            kw.update({ \
                'UnitID'       : self.unit_id
            })
        if self.illustration:
            kw.update({ \
                'Illustration' : '%s%s' % (_globals['PRODUCT_IMAGES_PATH'], self.illustration)
            })
        if self.style:
            kw.update({ \
                'Style'        : self.style
            })
        if self.icon:
            kw.update({ \
                'Icon'         : self.icon
            })
        return make_xmltag(self.id, indent=indent or self.default_xml_indent, encoding=encoding, **kw)

    def getContent(self, indent=None, encoding=None, locale=None):
        if not self.isXMLOutput():
            return None
        kw = {}
        if self.kind == 'DISPLAY FIELD' and self.type in ('ALERT', 'IMAGE', 'PICTURE',):
            if self.type == 'ALERT':
                kw.update({ \
                    'CurrentValue' : localize(ALERT_CONTENT % self.title, locale),
                })
            if self.type == 'IMAGE':
                if ':' in self.title:
                    label, value = self.title.split(':')
                    value = value.replace(PRODUCT_LOCATION, locale or DEFAULT_LOCALE)
                else:
                    label, value = '', self.title
                kw.update({ \
                    'Label'        : label,
                    'CurrentValue' : value,
                })
            if self.type == 'PICTURE':
                kw.update({ \
                    'CurrentValue' : self.title.replace(PRODUCT_LOCATION, locale or DEFAULT_LOCALE),
                })
        elif self.kind in ('DISPLAY FIELD',):
            kw.update({ \
                'CurrentValue' : localize(self.title, locale),
                'Description'  : localize(self.note or (self.parent is not None and self.parent.note) or '', locale),
            })
        else:
            kw.update({ \
                'Label'        : localize(self.title, locale),
                'Description'  : localize(self.note or (self.parent is not None and self.parent.note) or '', locale),
            })
        return make_xmltag(self.id, indent=indent or self.default_xml_indent, encoding=encoding, **kw)

class ProductInfoTag(ProductAbstract):
    """
        Инфотег
    """
    metatype = tT
    
    def __init__(self, id, title=None):
        super(ProductInfoTag, self).__init__(id, title)

    def isValid(self):
        return self.id and True or False

    def isXMLOutput(self):
        return True # self.title and True or False

    def group(self):
        return self.parent.parent.parent is not None and self.parent.parent.parent.id or ''

    def subgroup(self):
        return self.parent.parent is not None and self.parent.parent.id or ''

    def get_html_title(self, locale=None):
        title = self.parent.title or self.parent.objects[self.parent.ids[0]].title
        return '%s' % (localize(title, locale) or '')

    def getForm(self, indent=None, encoding=None, locale=None):
        if not self.isXMLOutput():
            return None
        kw = {}
        inline = False
        if self.kind in ('CONSTANT',):
            kw.update({ \
                'Kind'         : self.kind,
                'Type'         : self.type,
                #'GroupID'      : self.group(),
                #'SubGroupID'   : self.subgroup(),
            })
            group_type = 'SubGroupID'
            group_value = ''
            if self.parent.parent is not None:
                group_type = self.parent.parent.metatype == tG and 'GroupID' or 'SubGroupID'
                group_value = self.parent.parent.id
            kw.update({ \
                group_type     : group_value
            })
            inline = True
        elif self.kind in ('INPUT FIELD',):
            kw.update({ \
                'Kind'         : self.kind,
                'Type'         : self.type,
                'GroupID'      : self.group(),
                'SubGroupID'   : self.subgroup(),
            })
        elif self.kind in ('SIMPLE INPUT',):
            kw.update({ \
                'Kind'         : self.isOverridable() and 'INPUT FIELD' or self.kind,
                'Type'         : self.type,
                'GroupID'      : self.group(),
            })
        elif self.kind in ('POP-UP MENU', 'COMBOBOX', 'SELECT', 'H-LIST', 'H-LIST-BOX',):
            kw.update({ \
                'Kind'         : self.kind,
                'Type'         : self.type,
                'GroupID'      : self.group(),
                'SubGroupID'   : self.subgroup(),
            })
        if self.id_1c:
            kw.update({ \
                'ID'           : self.id_1c 
            })
        if self.format:
            kw.update({ \
                'Format'       : self.format
            })
        if self.current_value is not None:
            kw.update({ \
                'CurrentValue' : self.current_value or '',
            })
        if self.unit_id:
            kw.update({ \
                'UnitID'       : self.unit_id
            })
        if self.style:
            kw.update({ \
                'Style'        : self.style
            })
        if self.icon:
            kw.update({ \
                'Icon'         : self.icon
            })
        return make_xmltag(self.id, indent=indent or self.default_xml_indent, inline=inline, encoding=encoding, **kw)

class ProductRequirement(ProductAbstract):
    """
        Требование
    """
    metatype = tR
    
    def __init__(self, id, title=None):
        super(ProductRequirement, self).__init__(id, title)

    def isValid(self):
        return self.description and self.script and True or False

class ProductNote(ProductAbstract):
    """
        Подсказка
    """
    metatype = tN
    
    def __init__(self, id, title=None):
        super(ProductNote, self).__init__(id, title)

    def isValid(self):
        return self.note and True or False

class ProductMargin(ProductAbstract):
    """
        Наценка
    """
    metatype = tM
    
    def __init__(self, id, title=None):
        super(ProductMargin, self).__init__(id, title)

    def isValid(self):
        return self.num and self.script and True or False

class ProductConstant(ProductAbstract):
    """
        Константа
    """
    metatype = tC
    
    def __init__(self, id, title=None):
        super(ProductConstant, self).__init__(id, title)
        self.kind = 'CONSTANT'
        self.type = None
        self.value = None

    def _init_state(self, data):
        if not data:
            return

        if data.has_key('id') and ':' in data['id']:
            self.id, self.attr_id = data['id'].split(':')
            product_declarations.add(self.id)
        if data.has_key('num') and data['num']:
            self.num = data['num']
        if data.has_key('kind') and data['kind']:
            if upper(data['kind']) in VALID_CONSTANT_KINDS:
                self.kind = data['kind']
        if data.has_key('type') and data['type']:
            if upper(data['type']) in VALID_CONSTANT_TYPES:
                self.type = data['type']
                self._init_type()
        if (data.has_key('value1') and data['value1']) or (data.has_key('value2') and data['value2']):
            if self.type == 'LIST':
                pass
            else:
                if data['value1'] and data['value2']:
                    value = data['value1'], data['value2']
                else:
                    value = data['value1'] or data['value2']
                self.set_value(value)
        if data.has_key('description') and data['description']:
            self.description = data['description']

    def _init_type(self):
        if self.type == 'LIST':
            self.value = []

    def set_value(self, value):
        if self.type == 'LIST':
            self.add_listitem(value)
        elif self.type in ('STRING', 'MSG',):
            self.value = unicode(value)
        elif self.type == 'NUMBER':
            self.value = value
        elif self.type == 'BOOLEAN':
            self.value = value and True or False

    def get_value(self, indent='', locale=None):
        if not self.type == 'LIST':
            return localize(self.value, locale=locale)
        value = u'%s0%s' % (indent + ' '*8, _globals['LIST_INDEX_SEPARATOR'])
        for index in range(self.size()):
            value += u'%s%s%s%s' % (EOL, indent + ' '*8, self.get_listitem(index, locale=locale), _globals['LIST_VALUE_SEPARATOR'])
        if self.size() > 0:
            value = value[:-1]
        return u'%s%s%s%s' % (EOL, value, EOL, indent + ' '*4)

    def add_listitem(self, value):
        if type(value) in SequenceTypes:
            item = [unicode(x) for x in value if x]
        else:
            item = unicode(value)
        self.value.append(item)

    def get_listitem(self, index, separator=None, locale=None):
        if separator is None:
            separator = _globals['LIST_ITEM_SEPARATOR']
        if index < 0  or index > len(self.value) - 1:
            return None
        if type(self.value[index]) in SequenceTypes and locale:
            s = ' '
            value = []
            for v in self.value[index]:
                x = localize(v, locale)
                if v != x:
                    value.append(x)
                else:
                    value.append(s.join([localize(x, locale) for x in v.split(s)]))
            return separator.join(value)
        else:
            return separator.join([localize(x, locale) for x in self.value[index]])

    def size(self):
        if self.type == 'LIST':
            x = len(self.value)
        elif self.type == 'STRING':
            x = len(self.value)
        elif self.type == 'NUMBER':
            x = self.__sizeof__()
        elif self.type == 'BOOLEAN':
            x = 1
        return x

    def isValid(self):
        return self.id and self.type and self.value is not None and True or False

    def getForm(self, indent=None, encoding=None, locale=None):
        if not self.isXMLOutput():
            return None
        kw = {}
        inline = True
        if self.kind == 'CONSTANT' and self.type in VALID_TYPE_VALUES:
            kw.update({ \
                'Kind'         : self.kind,
                'Type'         : self.type,
            })
        if self.kind == 'PROGRESS':
            kw.update({ \
                'Kind'         : self.kind,
                'Type'         : self.type,
            })
        if self.kind == 'PROGRESS' and getattr(self, 'attr_id', None) is not None:
            kw.update({ \
                'ID'           : self.attr_id,
            })
            inline = False
        if self.value is not None:
            kw.update({ \
                'CurrentValue' : self.get_value(indent, locale=locale),
            })
            inline = False
        return make_xmltag(self.id, indent=indent or self.default_xml_indent, inline=inline, encoding=encoding, **kw)

    def get_html_value(self, locale=None):
        # Locale depended types
        if self.type in ('ALERT', 'IMAGE',):
            return "''"
        # and other
        if self.type == 'BOOLEAN':
            return '%s' % (self.value and 'true' or 'false')
        elif self.type == 'NUMBER':
            return '%s' % str(self.value or 0)
        elif self.type in ('STRING',):
            return "'%s'" % (self.value or '')
        elif self.type in ('MSG',):
            return "'%s'" % self.get_value(locale=locale)
        elif self.type == 'LIST':
            return "'%s'" % re.sub(r'\n[\s]+', '', self.get_value())
        else:
            return "null"

class ProductTerm(object):
    """
        Словарный термин
    """
    attr_id = 'num'
    default_value = None
    
    def __init__(self, attr=None):
        try:
            self.id = self.attr_id in attr.keys() and int(attr[self.attr_id])
        except:
            self.id = None
        self.key = None
        self.values = {}
        for locale in LOCALES:
            value = attr[locale]
            if locale == DEFAULT_LOCALE:
                self.key = unCapitalize(value)
            elif locale not in self.values.keys() or not self.values[locale]:
                self.values[locale] = unCapitalize(value)

    def __repr__(self):
        return self.key.encode(default_unicode) or '' # default_encoding

    def isValid(self):
        return self.id and self.key and 1 or 0

    def get_key(self):
        return self.key or u''

    def get_default_value(self):
        return self.get_key() or self.default_value

    def get_synonym(self):
        return LOCALE_ENG in self.values.keys() and self.values[LOCALE_ENG]

    def get_value(self, locale=None):
        if not (locale and self.isValid()):
            value = self.get_default_value()
        else:
            value = \
                locale in self.values.keys() and self.values[locale] or \
                self.get_synonym() or \
                self.get_default_value()
        return re.sub(r'\'', '', value)

class ProductVersion(object):
    """
        Версия продукта (запись журнала изменений)
    """
    attr_id = 'version'

    def __init__(self, attr={}):
        try:
            self.id = self.attr_id in attr.keys() and attr[self.attr_id]
        except:
            self.id = None
        self.date = attr.get('date') or 0
        self.description = attr.get('description')
        self.files = attr.get('files')
        self.is_web = attr.get('web') and True or False

    def __repr__(self):
        return self.description.encode(default_unicode) or '' # default_encoding

    def get_version(self):
        return self.id

    def get_date(self):
        x = get_datetime(self.date)
        return x and x.strftime(OUTPUT_DATE_FORMAT) or ''

    def get_description(self):
        s = re.sub(r'\n', '<br>', self.description or u'')
        return s

    def get_log_record(self):
        return '%s\t%s\t%s' % (self.get_date(), self.get_version() or ' ', self.get_description())


##  =====================================================================================================================  ##

class Logger():
    
    def __init__(self, to_file=None, encoding=default_unicode, mode='w+', bom=True, codecs=True):
        self.is_to_file = to_file and 1 or 0
        self.encoding = encoding
        self.fo = None
        if IsDisableOutput and to_file:
            pass
        elif to_file:
            self.to_file = to_file
            if codecs:
                self.codecs_open(mode, bom)
            else:
                self.open(mode)
            print '--> %s' % to_file
        else:
            #self.set_default_encoding()
            pass

    def codecs_open(self, mode, bom=False):
        self.fo = codecs.open(self.to_file, encoding=self.encoding, mode=mode)
        if bom:
            self.fo.write(codecs.BOM_UTF8.decode(self.encoding))

    def open(self, mode):
        self.fo = open(self.to_file, mode=mode)

    def set_default_encoding(self, encoding=default_unicode):
        if sys.getdefaultencoding() == 'ascii':
            reload(sys)
            sys.setdefaultencoding(encoding)
        print '--> %s' % sys.getdefaultencoding()

    def out(self, line, eol=None):
        if not line:
            return
        elif not (self.fo or self.is_to_file):
            try:
                print '--> %s' % line
            except:
                if type(line) is UnicodeType:
                    v = ''
                    for x in line:
                        try:
                            print x,
                            v += x.encode(default_encoding, 'ignore')
                        except:
                            v += '?'
                    print '' #'%s==> Unicode ERROR %s' % (EOL, type(line))
                else:
                    print '--> %s' % line.decode(default_encoding, 'ignore')
        elif IsDisableOutput:
            return
        else:
            if type(line) in StringTypes:
                try:
                    self.fo.write(line)
                except:
                    try:
                        self.fo.write(unicode(line, self.encoding))
                    except:
                        #print line.decode(default_encoding)
                        #raise
                        try:
                            self.fo.write(line.decode(default_encoding)) #, 'replace'
                            #print line
                        except:
                            #self.fo.write('xxx')
                            raise
                if not eol:
                    eol = EOL
                if not line == eol:
                    self.fo.write(eol)

    def readlines(self):
        return self.fo.readlines()

    def update_and_close(self, lines):
        self.fo.close()
        self.open('w')
        for line in lines:
            self.fo.write(line+EOL)
        self.close()

    def close(self):
        if IsDisableOutput:
            return
        if not self.fo:
            return
        self.fo.close()

##  =====================================================================================================================  ##

def get_lazy_id(metatype):
    if not metatype in _globals:
        _globals[metatype] = 0
    x = '%s%s' % (metatype[0:1].upper(), _globals[metatype])
    _globals[metatype] += 1
    return x

def get_cell_value(value):
    if not value and type(value) in StringTypes:
        return u''
    if type(value) is StringType:
        value = unicode(value.strip(), default_unicode)
    elif type(value) is UnicodeType:
        value = value.strip()
    return value

def get_datetime(serial):
    if not serial:
        return None
    seconds = (serial - 25569) * 86400.0
    return datetime.datetime.utcfromtimestamp(seconds)

def create_instance(index, attr, row, container=None):
    """
        Создать экземпляр описания
    """
    ob = None
    if index == 1:
        if attr['id'] and not attr['num'] and attr['group']:
            ob = ProductGroup(attr['id'])
        elif attr['id'] and attr['num'] and attr['parameter']:
            ob = ProductParameter(attr['id'])
        elif attr['id'] and not (attr['num'] or attr['parameter']) and attr['value']:
            ob = ProductValue(attr['id'])
        elif not (attr['num'] or attr['parameter'] or attr['value']) and attr['id']:
            ob = ProductInfoTag(attr['id'])
        elif attr['parameter'] and attr['parameter'].strip() == '!' and attr['value'] and attr['script']:
            ob = ProductRequirement(get_lazy_id('rid'))
            ob.set_description(attr['value'])
        elif not (attr['num'] or attr['value'] or attr['id']) and attr['parameter']:
            ob = ProductNote(get_lazy_id('nid'))
            ob.set_note(attr['parameter'])
    elif index == 2:
        if attr['num'] and attr['parameter'] and attr['script']:
            ob = ProductMargin(get_lazy_id('mid'))
    elif index == 3:
        if attr['num'] and attr['id'] and attr['type']:
            ob = ProductConstant(attr['id'])
        elif attr['id'] and container is not None:
            container.set_value((attr['id'], attr['value1'], attr['value2']))
            return container
    elif index == 5:
        ob = ProductTerm(attr)

    if hasattr(ob, 'metatype') and ob is not None:
        # номер строки описания
        ob.set_line(row)
        # инициализировать поля описания
        ob._init_state(attr)

    return ob is not None and ob or None

def _product_attrs_definition(name, index=1):
    """
        Returns product's attributes definition objects list
        ----------------------------------------------------
        Sheet 1
    """
    book = xlrd.open_workbook(name)
    sheet = book.sheet_by_index(index)

    attrs = []

    if index == 1:
        assert sheet.ncols >= 20, "Sheet %s has invalid columns number: %s" % (index, sheet.ncols)
        defs = S1_COLUMNS_LINE
        start_row = 5
    else:
        start_row = 0

    group_ob = None
    parameter_ob = None
    value_ob = None

    for row in range(start_row, sheet.nrows):
        IsBreak = False
        IsTrace = False
        IsHide = False

        attr = {}

        for column in range(sheet.ncols):
            value = sheet.cell_value(row, column) # or None

            if column == 0 and value:
                if value in OP_RECORD_DELETE:
                    IsBreak = True
                    break
                if value in OP_IS_TRACE:
                    IsTrace = True
                if value in OP_IS_HIDE:
                    IsHide = True

            value = get_cell_value(value)

            if not defs[column]:
                continue

            attr[defs[column]] = value

        if IsBreak:
            continue

        ob = create_instance(index, attr, row)

        if ob is None or not repr(ob):
            continue

        if IsTrace:
            ob.is_trace = True

        if IsHide:
            ob.set_status('hide')

        if ob.get_type() == tG:
            group_ob = ob
            parameter_ob = None
            value_ob = None
            
        elif ob.get_type() == tP:
            parameter_ob = ob
            value_ob = None
            if group_ob is not None:
                group_ob.set(ob)
            ob.set_parent([group_ob])
            
        elif ob.get_type() == tV:
            value_ob = ob
            if group_ob is not None:
                group_ob.set(ob)
            if parameter_ob is not None:
                parameter_ob.set(ob)
            ob.set_parent([parameter_ob, group_ob])
            
        elif ob.get_type() == tT:
            if group_ob is not None:
                group_ob.set(ob)
            if parameter_ob is not None:
                parameter_ob.set(ob)
            if value_ob is not None:
                value_ob.set(ob)
            ob.set_parent([value_ob, parameter_ob, group_ob])
            
        elif ob.get_type() == tR:
            if group_ob is not None:
                group_ob.set(ob)
            if parameter_ob is not None:
                parameter_ob.set(ob)
            ob.set_parent([parameter_ob, group_ob])
            
        elif ob.get_type() == tN:
            if value_ob is not None:
                setattr(value_ob, 'note', ob.note)
            if parameter_ob is not None:
                setattr(parameter_ob, 'note', ob.note)

        if ob.isValid():
            attrs.append(ob)

    del sheet
    del book

    return attrs

def _product_margins_definition(name, index=2):
    """
        Returns pricing's attributes definition objects list
        ----------------------------------------------------
        Sheet 2
    """
    book = xlrd.open_workbook(name)
    sheet = book.sheet_by_index(index)

    attrs = []

    if index == 2:
        assert sheet.ncols == 8, "Sheet %s has invalid columns number: %s" % (index, sheet.ncols)
        defs = S2_COLUMNS_LINE
        start_row = 4
    else:
        start_row = 0

    for row in range(start_row, sheet.nrows):
        IsBreak = False
        attr = {}

        for column in range(sheet.ncols):
            value = sheet.cell_value(row, column) # or None

            if column == 0 and value:
                if value in OP_RECORD_DELETE:
                    IsBreak = True
                    break

            value = get_cell_value(value)

            if not defs[column]:
                continue
            
            attr[defs[column]] = value

        if IsBreak:
            continue

        ob = create_instance(index, attr, row)

        if ob is None or not repr(ob):
            continue

        if ob.isValid():
            attrs.append(ob)

    del sheet
    del book

    return attrs

def _product_constants_definition(name, index=3):
    """
        Returns constants attributes definition objects list
        ----------------------------------------------------
        Sheet 3
    """
    book = xlrd.open_workbook(name)
    sheet = book.sheet_by_index(index)

    attrs = []

    if index == 3:
        assert sheet.ncols == 8, "Sheet %s has invalid columns number: %s" % (index, sheet.ncols)
        defs = S3_COLUMNS_LINE
        start_row = 4
    else:
        start_row = 0

    container = None
    IsHide = False

    for row in range(start_row, sheet.nrows):
        IsBreak = False
        attr = {}

        for column in range(sheet.ncols):
            value = sheet.cell_value(row, column) # or None

            if IsHide:
                if column == 0:
                    continue
                elif column == 1 and value:
                    IsHide = False
                else:
                    IsBreak = True
                    break

            if column == 0 and value:
                if value in OP_RECORD_DELETE:
                    IsBreak = True
                    break
                if value in OP_IS_HIDE:
                    IsBreak = True
                    IsHide = True
                    break

            value = get_cell_value(value)

            if not defs[column]:
                continue
            
            attr[defs[column]] = value

        if IsBreak:
            continue

        ob = create_instance(index, attr, row, container=container)

        if ob is None or not repr(ob):
            continue

        if ob not in attrs:
            attrs.append(ob)

        container = ob

    del sheet
    del book

    return attrs

def _product_locale_definition(name, index=5, attrs={}):
    """
        Returns product_locale attributes definition objects list
        ----------------------------------------------------
        Sheet 5
    """
    book = xlrd.open_workbook(name)
    sheet = book.sheet_by_index(index)

    if not index or index == 5:
        assert sheet.ncols == len(S5_COLUMNS_LINE), "Sheet %s has invalid columns number: %s" % (index, sheet.ncols)
        defs = S5_COLUMNS_LINE
        start_row = 4
    else:
        start_row = 0

    for row in range(start_row, sheet.nrows):
        IsBreak = False
        attr = {}

        for column in range(sheet.ncols):
            value = sheet.cell_value(row, column) # or None

            if column == 0 and value:
                if value in OP_RECORD_DELETE:
                    IsBreak = True
                    break

            value = get_cell_value(value)

            if not defs[column]:
                continue

            if defs[column] == DEFAULT_LOCALE:
                if not value:
                    IsBreak = True
                    break
                #value = value.encode(default_encoding, 'ignore').strip()

            attr[defs[column]] = value

        if IsBreak:
            continue

        ob = create_instance(5, attr, row)

        if ob is None or not repr(ob):
            continue

        attrs[ob.get_key()] = ob

    del sheet
    del book

    return attrs

def _product_version_definition(name, index=6):
    """
        Returns current configurator version number
        -------------------------------------------
        Sheet 6
    """
    book = xlrd.open_workbook(name)
    sheet = book.sheet_by_index(index)

    attrs = []

    if index == 6:
        assert sheet.ncols == 5, "Sheet %s has invalid columns number: %s" % (index, sheet.ncols)
        defs = S6_COLUMNS_LINE
        start_row = 4
    else:
        start_row = 0

    container = None
    IsHide = False

    for row in range(start_row, sheet.nrows):
        IsBreak = False
        IsWeb = False
        attr = {}

        for column in range(sheet.ncols):
            value = sheet.cell_value(row, column) # or None

            if column == 0 and value:
                if value in OP_RECORD_DELETE:
                    IsBreak = True
                    break
                if value in OP_WEB:
                    IsWeb = True

            value = get_cell_value(value)

            if not defs[column]:
                continue
            
            attr[defs[column]] = value

        attr['web'] = IsWeb

        if IsBreak:
            continue

        ob = ProductVersion(attr)

        if ob is None or not repr(ob):
            continue

        attrs.append(ob)

    del sheet
    del book

    return attrs

def valid_show_option(ob):
    IsValid = 1
    if IsShowOnlyGroups and not ob.get_type() == tG:
        IsValid = 0
    if IsShowOnlyParameters and not ob.get_type() == tP:
        IsValid = 0
    if IsShowOnlyValues and not ob.get_type() == tV:
        IsValid = 0
    if IsShowOnlyInfoTags and not ob.get_type() == tT:
        IsValid = 0
    if IsShowOnlyRequirements and not ob.get_type() == tR:
        IsValid = 0
    if IsShowOnlyMargins and not ob.get_type() == tM:
        IsValid = 0
    if IsShowOnlyConstants and not ob.get_type() == tC:
        IsValid = 0
    return IsValid

def content_generator(attrs, margins, constants, versions, params={}, is_trace=None):
    """
        Generates and prints into output files content in accodance with given params.

        Attributes:

            attrs -- product attributes collection, list
            
            margins -- pricing collection, list
            
            constants -- constants collection, list
            
            versions -- changes log, list
            
            params -- dict:
            
                'form' : form file name
                
                'content' : content file name
                
                'script' : script file name
                
                'encoding' : output files encoding, by default - default_unicode

            is_trace -- auto trace mode (true/false).
    """
    IsForm = params['form'] and True or False
    IsContent = params['content'] and True or False
    IsScript = params['script'] and True or False
    IsJavaScript = params['js'] and True or False
    IsLog = params['log'] and True or False
    
    encoding = params.has_key('encoding') and params['encoding'] or default_unicode
    level1 = ' '*4
    level2 = ' '*8

    def clean(s):
        return re.sub(r'%s{2,}' % EOL, EOL*2, re.sub(r'\s*else\s*\{\s*\}%s*(?si)' % EOL, EOL*2, s)) or s

    def sout(s, no_clean=None):
        # перекодировка (вывод полей) в default_unicode
        if no_clean:
            return s or ''
        return s and clean(s) or ''

    def chdir(s):
        if not os.path.exists(s):
            os.mkdir(s)

    global _globals, short_version, default_path
    short_version = sout(short_version)

    # Контроль идентификатора (id) и маршрута к рабочим директориям (path)
    path = id = ''
    if _globals['PRODUCT_PATH']:
        if ',' in _globals['PRODUCT_PATH']:
            path, id = _globals['PRODUCT_PATH'].split(',')
            id = id.strip()
        else:
            path = _globals['PRODUCT_PATH']

        # id должен быть задан всегда
        assert id, "Invalid object id: [%s]" % id

        if path.endswith(':'):
            path = os.path.join(path, '\\')
            print '--> path=%s' % path

    # Проверка/создание структуры директорий
    if not os.path.exists(path):
        path = '.'
    chdir(os.path.join(path, default_path))
    chdir(os.path.join(path, default_path, id))

    _globals['PRODUCT_IMAGES_PATH'] = '/%s/%s/images/' % (default_path, id)
    
    _globals['version'] = {
        'name' : 'Script',
        'description' : 'Generated by Helpgen.py, %s' % short_version,
        'date' : 'Date: %s' % date_version
    }

    constant_obs = {}
    for constant in constants:
        if constant.get_type() == tC and constant.type != 'MSG':
            if constant.value is not None:
                constant_obs[constant.id] = constant

    # -------------------------
    #   Файл формы (form.xml)
    # -------------------------

    if IsForm:
        # Header
        form_file = os.path.join(path, default_path, id, params['form'])
        fo = Logger(form_file, encoding)
        fo.out(FORM_TEMPLATE['header'])
        fo.out('<!-- %s -->' % sout(_globals['PRODUCT_TITLE']))
        fo.out('<!-- %s -->' % _globals['version']['description'])
        fo.out('<!-- %s -->' % _globals['version']['date'])
        fo.out(FORM_TEMPLATE['doctype'])
        fo.out(FORM_TEMPLATE['document_header'] % {'id':_globals['PRODUCT_ID'], 'lang':'RUS'})

        # Body (Параметры изделия)
        fo.out('%s<!-- %s -->' % (level2, sout('РАЗДЕЛ ДЛЯ ПОЛЕЙ ФОРМЫ КОНФИГУРАТОРА')))
        group = None
        parameter = None
        
        for attr in attrs:
            if attr.isHide():
                continue
            if attr.get_type() == tG:
                if group is not None:
                    fo.out(EOL)
                if attr.title:
                    fo.out('%s<!-- %s -->' % (level2, sout(attr.title)))
                fo.out( \
                    sout(attr.getForm(level2), no_clean=1)
                )
                group = attr
                parameter = None
            elif attr.get_type() == tP:
                if parameter is not None:
                    fo.out(EOL)
                if attr.title:
                    fo.out('%s<!-- %s -->' % (level2, sout(attr.title)))
                parameter = attr
            elif attr.get_type() in (tV, tT,):
                if attr.title:
                    fo.out('%s<!-- %s -->' % (level2, sout(attr.title)))
                elif attr.get_type() == tT and attr.id_1c:
                    fo.out('%s<!-- %s -->' % (level2, sout('связь с 1С')))
                fo.out( \
                    sout(attr.getForm(level2), no_clean=1)
                )

        fo.out(EOL)
        fo.out('%s<!-- %s -->' % (level2, sout('РАЗДЕЛ КОНСТАНТ')))

        # Body (Константы)
        for constant in constants:
            if constant.get_type() == tC and not constant.type in LOCALIZED_CONSTANT_TYPES:
                fo.out( \
                    sout(constant.getForm(level2))
                )
                fo.out(EOL)

        # Footer
        fo.out(FORM_TEMPLATE['document_footer'] % {'id':_globals['PRODUCT_ID']})
        fo.close()
        
        del fo

    # --------------------------------------------
    #   Файлы языковой локализации (content.xml)
    # --------------------------------------------

    if IsContent:
        for locale in LOCALES:
            # Проверка/создание структуры директорий
            chdir(os.path.join(path, locale))
            chdir(os.path.join(path, locale, default_path))
            chdir(os.path.join(path, locale, default_path, id))

            # Header
            content_file = os.path.join(path, locale, default_path, id, params['content'])
            co = Logger(content_file, encoding)
            co.out(CONTENT_TEMPLATE['header'])
            co.out('<!-- %s -->' % sout(_globals['PRODUCT_TITLE']))
            co.out('<!-- %s -->' % _globals['version']['description'])
            co.out('<!-- %s -->' % _globals['version']['date'])
            co.out(CONTENT_TEMPLATE['doctype'])
            co.out(CONTENT_TEMPLATE['document_header'] % {'id':_globals['PRODUCT_ID'], 'lang':upper(locale)})

            # Body (Документация)
            co.out('%s%s' % (level1, "<table name='KnowledgeRepository'>"))
            co.out(make_xmltag('c0', indent=level2, 
                image='icon1.png', 
                description=localize(u'Конфигуратор изделий (WEB-помощник). Руководство пользователя.', locale),
                location='documents/HelperInfo.pdf'))

            if _globals['PRODUCT_INFO']: 
                n = 1
                for image, description, location in _globals['PRODUCT_INFO']:
                    co.out( \
                        make_xmltag('c%s' % n, indent=level2, image=image, description=localize(description.decode(default_encoding), locale), 
                            location=location.replace(PRODUCT_LOCATION, locale)
                        )
                    )
                    n += 1

            co.out('%s%s%s' % (level1, "</table>", EOL))
            co.out('%s<!-- %s -->' % (level1, sout('ОПИСАНИЕ ГРУПП')))
            co.out("%s<table name='GroupAndSubgroup'>" % level1)

            # Body (Группы)
            for attr in attrs:
                if attr.isHide():
                    continue
                if attr.get_type() in (tG, tP,):
                    co.out( \
                        sout(attr.getContent(level2, locale=locale))
                    )

            co.out('%s</table>' % level1)
            co.out(EOL)
            co.out('%s<!-- %s -->' % (level1, sout('РАЗДЕЛ ДЛЯ ПОЛЕЙ ФОРМЫ КОНФИГУРАТОРА')))
            co.out("%s<table name='HelperTaskDescription'>" % level1)

            # Body (Поля формы)
            for attr in attrs:
                if attr.isHide():
                    continue
                if attr.get_type() in (tV,):
                    co.out( \
                        sout(attr.getContent(level2, locale=locale))
                    )

            co.out(EOL)
            co.out('%s<!-- %s -->' % (level2, sout('РАЗДЕЛ КОНСТАНТ')))

            # Body (Константы)
            for constant in constants:
                if constant.get_type() == tC and constant.type in LOCALIZED_CONSTANT_TYPES:
                    co.out( \
                        sout(replaceQuotedValues(constant.getForm(level2, locale=locale), constant_obs))
                    )
                    co.out(EOL)

            co.out('%s</table>' % level1)

            # Footer
            co.out(CONTENT_TEMPLATE['document_footer'])
            co.close()
            
            del co

    # -----------------------------
    #   Файл скрипта (script.txt)
    # -----------------------------

    if IsScript:
        # Header
        script_file = os.path.join(path, default_path, id, params['script'])
        so = Logger(script_file, encoding)
        so.out('// %s' % sout(_globals['PRODUCT_ID']))
        so.out('// %s' % ('-'*len(_globals['PRODUCT_ID'])))
        so.out('// %s' % _globals['version']['description'])
        so.out('// %s' % _globals['version']['date'])

        so.out(EOL)
        so.out("wizardID = '%s';" % _globals['PRODUCT_ID'])

        # Body (Сообщения автотрассировки)
        trace = {'on' : 'false', 'messages' : '', 'code' : ''}
        traced_ids = []

        # Body (Скрипты)
        for section in _globals['SCRIPT_SECTIONS_SORTED']:
            template = SCRIPT_TEMPLATE[section]
            code = ''
            code_then = ''
            code_else = ''
            code_extra = ''
            code_trace = ''

            for attr in attrs + margins:
                if attr.isHide():
                    continue
                attr_id = attr.id

                if attr.get_type() in (tG, tP, tV, tT, tR, tM,):
                    script = attr.getScript(section)

                    if attr.isTrace() and attr_id not in traced_ids and section != 'End':
                        trace['on'] = 'true'
                        trace['messages'] += '%s%s"==> %s [" + %s + "]<br>" +' % (EOL, level2, attr_id, attr_id)
                        traced_ids.append(attr_id)

                    if not script:
                        continue

                    for key, value in script:
                        #value = value.encode(default_encoding) #, 'ignore'
                        if key == 'then':
                            code_then += code_then and EOL*2 + value or value
                        elif key == 'else':
                            code_else += code_else and EOL*2 + value or value
                        elif key == 'extra':
                            code_extra += code_extra and EOL*2 + value or value
                        elif key == 'trace':
                            code_trace += code_trace and EOL*2 + value or value
                        else:
                            code += code and EOL*2 + value or value

                    if code_trace:
                        trace['code'] += code_trace

            if section == 'Trace':
                so.out(sout(template % trace))
            else:
                x = {'code' : code, 'code_then' : code_then, 'code_else' : code_else, 'code_extra' : code_extra}
                x.update(_globals['OPTIONS'])
                so.out(sout(template % x))

        so.out(EOL)
        so.out("NoticeMessage += (NoticeMessage ? '<br>' : '') + '%s<br>%s<br>%s<br>%s'" % (
            '* '*3, _globals['version']['description'], _globals['PRODUCT_RELEASE'], _globals['version']['date'])
        )

        # Footer
        so.close()
        
        del so

    # ---------------------------------
    #   Файл Java-скрипта (script.js)
    # ---------------------------------

    if IsScript and IsJavaScript:
        # Header
        script_file = os.path.join(path, default_path, id, params['js'])
        so = Logger(script_file, encoding)
        so.out('// %s' % sout(_globals['PRODUCT_ID']))
        so.out('// %s' % ('-'*len(_globals['PRODUCT_ID'])))
        so.out('// %s' % _globals['version']['description'])
        so.out('// %s' % _globals['version']['date'])

        so.out(EOL)

        source = file(os.path.join(path, default_path, id, params['script']), 'rb')
        b = source.read()
        source.close()

        b = re.sub(r'\/\*(.*?)\*\/', '', b)
        b = re.sub(r'\/\/(.*?)\n', '\n', b)
        b = re.sub(r'(var)\s+([\w]+)(:[\w]+)\s*=\s*', r'\1 \2 = ', b)
        b = re.sub(r'(var)\s+([\w]+)(:[\w]+);', r'\1 \2;', b)
        b = re.sub(r'\n{3,}', r'\n\n', b)

        r = re.compile(r'(function\s+.*)\((.*)\)([\:\w\s\n]*?)\{')
        m = r.search(b)
        while m:
            if m.group(3):
                x = re.sub(r'(\w*)\:(\w*)(,?)', r'\1\3', m.group(2))
                b = b[:m.start()] + m.group(1) + '(' + x + ') {' + b[m.end():]
            m = r.search(b, m.end())

        b = re.sub(r'(for)\s+(each)\s*(\()', r'\1 \3', b)
        b = re.sub(r'(new)\s+(Boolean)\s*\(([\w]+)\)(?ims)', r'\3', b)
        b = re.sub(r'(helperErrorCode)\s*=\s*(\'\')', r'\1 = 0', b)
        b = re.sub(r',\s+(\]|\}|\))', r'\1', b)

        if b:
            header = '' # XXX

            constant_template = JAVASCRIPT_TEMPLATE['Constant'].strip()
            field_template = JAVASCRIPT_TEMPLATE['Field'].strip()
            download_template = JAVASCRIPT_TEMPLATE['Download'].strip()
            upload_template = JAVASCRIPT_TEMPLATE['Upload'].strip()

            vars = []
            for constant in constants:
                if constant.get_type() == tC and constant.type not in LOCALIZED_CONSTANT_TYPES:
                    value = constant.get_html_value()
                    if value is not None:
                        vars.append(constant_template % {'id':constant.id, 'value':value})

            gids = []
            gattrs = []
            fields = []
            cis = []
            for attr in attrs:
                if attr.get_type() in (tV, tT,):
                    value = attr.get_html_value()
                    if value is not None:
                        if attr.isActive():
                            gids.append("'%s'" % attr.id)
                            gattrs.append("%s'%s':{'kind':'%s', 'type':'%s', 'subgroup':'%s'}" % (level1, attr.id, attr.kind, attr.type, attr.subgroup()))
                            if attr.id_1c:
                                cis.append("%s'%s':{'id':'%s', 'title':'%s'}" % (level1, attr.id, attr.id_1c, attr.get_html_title()))
                        x = {'id':attr.id, 'type':'gattrs.%s.type' % attr.id, 'value':value, 'indent':level1}
                        fields.append(field_template % x);

            brains = []
            footer = ''

            for n, line in enumerate(b.split(EOL)[1:]):
                brains.append('%s%s' % (level1, line))

            template = JAVASCRIPT_TEMPLATE['Body']
            x = { \
                'header'     : header, 
                'gids'       : ', '.join(gids), 
                'gattrs'     : EOI.join(gattrs),
                'cis'        : '', #EOI.join(cis),
                'vars'       : EOL.join(vars),
                'fields'     : EOL.join(fields), 
                'brains'     : EOL.join(brains).decode(default_unicode, 'ignore'), 
                'footer'     : footer
            }
            so.out(sout(template % x))

        so.close()
        del so

        make_min(os.path.join(path, default_path, id, 'script'))

    # --------------------------------------------
    #   Журнал изменений (log.txt)
    # --------------------------------------------

    if IsLog:
        log_file = os.path.join(path, default_path, id, params['log'])
        lo = Logger(log_file, encoding, bom=False)

        # Log (Версии)
        for ob in versions:
            if ob and ob.is_web:
                lo.out( \
                    sout(ob.get_log_record())
                )

        lo.close()
        del lo

    # ---------------------------------------
    #   Register module into startup config
    # ---------------------------------------

    if IsScript and IsJavaScript:
        config_file = os.path.join(path, model2_config)
        so = Logger(config_file, mode='a+', bom=False, codecs=False)

        pid = _globals['PRODUCT_ID']
        IsFound = False
        items = []

        for line in so.readlines():
            if not line.strip():
                continue
            x = line.split(':')
            id = x[0].strip()
            if id == pid:
                IsFound = True
            active = len(x) > 1 and x[1].strip() in ('1', '+', 'on',) and 1 or 0
            items.append((id, active))

        if not IsFound:
            items.append((pid, 1))

        so.update_and_close(['%-30s : %s' % (id, active) for id, active in sorted(items)])
        del so


if __name__ == "__main__":
    argv = sys.argv

    if len(argv) < 2 or argv[1].lower() in ('/h', '/help', '-h', 'help', '--help'):
        print '--> DoorHan Inc.'
        print '--> *Web-Helper* configurator\'s script parser.'
        print '--> '
        print '--> Format: helpgen.py [options] <xls-descriptor file name>'
        print '--> '
        print '--> Options:'
        print '--> '
        print '-->   -d:        IsDebug'
        print '-->   -D:        IsDeepDebug'
        print '-->   -s or S:   IsShowScriptDebug'
        print '-->   -A:        all debug options'
        print '-->   -F:        offline debug options (-dC+IsWithoutDB)'
        print '-->   -o<id>:    show object\'s container with given id'
        print '-->   -p<file>:  print output to file'
        print '-->   -g:        show only groups'
        print '-->   -p:        show only parameters'
        print '-->   -v:        show only values'
        print '-->   -t:        show only infotags'
        print '-->   -r:        show only requirements'
        print '-->   -m:        show only margins'
        print '-->   -c:        show only constants'
        print '-->   -n:        show notes'
        print '-->   -l:        show product_locale'
        print '-->   -C:        capitalize titles'
        print '-->   -U:        uncapitalize titles'
        print '--> '
        print '-->   --encoding=<default encoding for output>, default: %s' % default_encoding
        print '-->   --syntax   script syntax control (simple declarations only)'
        print '-->   --keep     disable files output'
        print '-->   --locale   print titles for translate'
        print '-->   --code     print source output'
        print '--> '
        print '--> Main argument:'
        print '--> '
        print '-->   <xls-descriptor file name> without extention, e.g.: product id'
        print '--> '
        print '--> %s[Python2]' % version
    else:
        for x in argv[1:]:
            if x.startswith('--'):
                if x[2:] == 'syntax':
                    IsSyntaxControl = 1
                elif x[2:] == 'keep':
                    IsDisableOutput = 1
                elif x[2:] == 'locale':
                    IsPrintLocale = 1
                elif x[2:] == 'code':
                    IsPrintCode = 1
                else:
                    r = re.compile('^--encoding=(.*)$')
                    m = r.search(x)
                    if m is not None and m.group():
                        default_encoding = m.group()
                continue
            if x.startswith('-'):
                if x[1:2] == 'o':
                    debug_object_id = x[2:]
                    assert debug_object_id, "Debug object id is not present!"
                    IsObjectDebug = True
                elif x[1:2] == 'p':
                    print_to_file = x[2:]
                    assert print_to_file, "File name to printing is not present!"
                    print_to_file = './' + x[2:]
                else:
                    if 'd' in x:
                        IsDebug = True
                    if 'D' in x:
                        IsDeepDebug = True
                    if 's' in x or 'S' in x:
                        IsShowScriptDebug = True
                    if 'A' in x:
                        IsDebug = True
                        IsDeepDebug = True
                        IsShowScriptDebug = True
                        IsShowNotes = True
                        IsShowLocations = True
                    if 'F' in x:
                        IsDebug = True
                        IsWithoutDB = True
                        IsCapitalizeTitles = True
                    if 'g' in x:
                        IsShowOnlyGroups = True
                    if 'p' in x:
                        IsShowOnlyParameters = True
                    if 'v' in x:
                        IsShowOnlyValues = True
                    if 't' in x:
                        IsShowOnlyInfoTags = True
                    if 'r' in x:
                        IsShowOnlyRequirements = True
                    if 'm' in x:
                        IsShowOnlyMargins = True
                    if 'c' in x:
                        IsShowOnlyConstants = True
                    if 'n' in x:
                        IsShowNotes = True
                    if 'l' in x:
                        IsShowLocations = True
                    if 'C' in x:
                        IsCapitalizeTitles = True
                    if 'U' in x:
                        IsUnCapitalizeTitles = True
            else:
                product = x

        assert product, "Product name is not present!"
        product_filename = '%s.xls' % product

        logger = Logger(print_to_file, mode=print_to_file and 'a+' or 'w+')

        product_info(product_filename, index=ID_LINE['sheet'])
        logger.out('-'*20)
        logger.out('%s: %s' % (_globals['PRODUCT_ID'], _globals['PRODUCT_TITLE'].decode(default_encoding, 'ignore')))

        if IsPrintLocale:
            locale_logger = Logger('locale.txt')
            locale_logger.out('%s' % _globals['PRODUCT_TITLE'].decode(default_encoding, 'ignore'))

        if IsPrintCode:
            code_logger = Logger('code.txt')
            code_logger.out('%s' % _globals['PRODUCT_TITLE'].decode(default_encoding, 'ignore'), EOR)

        ## ---------------------
        ##  Конструкция изделия
        ## ---------------------

        logger.out('product definitions')
        attrs = _product_attrs_definition(product_filename)

        logger.out('defs for %s attrs' % len(attrs))

        level1 = ' '*3
        level2 = ' '*6

        if IsDebug:
            for i, ob in enumerate(attrs):
                if IsObjectDebug and not (ob.id == debug_object_id or debug_object_id in ob.parents):
                    continue
                if not valid_show_option(ob):
                    continue
                logger.out('%03d %s [%s]: id=%s%s, line=%s' % (i, ob.get_type(), ob.get_status(), ob.id, ob.title and ', '+ob.title or '', ob.line))
                if IsDeepDebug:
                    if ob.get_type() in (tG, tP, tV, tT):
                        if ob.id_1c:
                            logger.out('%s ID 1C: %s' % (level2, ob.id_1c))
                        if ob.kind:
                            logger.out('%s Kind: %s' % (level2, ob.kind))
                        if ob.type:
                            logger.out('%s Type: %s' % (level2, ob.type))
                        if ob.format:
                            logger.out('%s Format: %s' % (level2, ob.format))
                        if ob.current_value is not None:
                            logger.out('%s CurrentValue: %s [%s]' % (level2, ob.current_value, type(ob.current_value)))
                        if ob.unit_id:
                            logger.out('%s UnitId: %s' % (level2, ob.unit_id))
                        if ob.parent:
                            logger.out('%s Parent: %s' % (level2, ob.parent.id))
                    if ob.get_type() == tR:
                        for line in ob.description.split(EOL):
                            try:
                                logger.out('%s %s' % (level2, line))
                            except:
                                logger.out('%s %s' % (level2, '...'))
                                #break
                if IsShowNotes and ob.get_type() == tP and ob.note:
                    try:
                        logger.out('%s Note: %s' % (level2, ob.note))
                    except:
                        logger.out('%s Note: %s' % (level2, '...'))
                if IsShowScriptDebug and ob.script:
                    for section in ob.script:
                        logger.out('%s S: %s' % (level1, section))
                        for key in ob.script[section]:
                            logger.out('%s S: %s' % (level1, key))
                            for n, line in enumerate(ob.script[section][key].split(EOL)):
                                logger.out('%s %03d %s' % (level2, n+1, line))

        ## -----------------
        ##  Ценообразование
        ## -----------------

        logger.out('pricing definitions')
        margins = _product_margins_definition(product_filename)

        logger.out('defs for %s margins' % len(margins))

        if IsDebug:
            for i, ob in enumerate(margins):
                if IsObjectDebug and not (ob.id == debug_object_id or debug_object_id in ob.parents):
                    continue
                if not valid_show_option(ob):
                    continue
                logger.out('%03d %s: id=%s%s, line=%s' % (i, ob.get_type(), ob.id, ob.title and ', '+ob.title or '', ob.line))
                if IsShowScriptDebug and ob.script:
                    for section in ob.script:
                        logger.out('%s S: %s' % (level1, section))
                        for key in ob.script[section]:
                            logger.out('%s S: %s' % (level1, key))
                            for n, line in enumerate(ob.script[section][key].split(EOL)):
                                logger.out('%s %03d %s' % (level2, n+1, line))

        ## -----------
        ##  Константы
        ## -----------

        logger.out('constants definitions')
        constants = _product_constants_definition(product_filename)

        logger.out('defs for %s constants' % len(constants))

        if IsDebug:
            for i, ob in enumerate(constants):
                if IsObjectDebug and not (ob.id == debug_object_id or debug_object_id in ob.parents):
                    continue
                if not valid_show_option(ob):
                    continue
                logger.out('%03d %s: id=%s, type=%s, line=%s' % (i, ob.get_type(), ob.id, ob.type, ob.line))
                if IsDeepDebug:
                    if ob.type == 'LIST':
                        for index in range(len(ob.value)):
                            logger.out('%s %s' % (level2, ob.get_listitem(index)))
                        logger.out('%s %s items' % (level2, ob.size()))
                    elif ob.type in ('STRING', 'MSG',):
                        logger.out('%s %s' % (level2, ob.value))
                    else:
                        logger.out('%s %s' % (level2, ob.value))

        ## ------------------
        ##  Словарь терминов
        ## ------------------

        logger.out('product locale definitions')
        product_locale = _product_locale_definition(default_product_locale, index=0)
        product_locale = _product_locale_definition(product_filename, attrs=product_locale)

        logger.out('defs for %s product locale: %s terms' % (len(LOCALES), len(product_locale)))

        if IsShowLocations and not IsObjectDebug:
            for i, key in enumerate(sorted(product_locale.keys())):
                ob = product_locale[key]
                if ob is not None and ob.isValid():
                    x = ob.get_key()
                    #try:
                    #    x = x.decode(default_encoding, 'ignore')
                    #except:
                    #    x = type(x) is UnicodeType and x.encode(default_unicode)
                    logger.out('%03d %s' % (ob.id, x))
                    #, ob.get_value(LOCALE_CZE).decode(default_encoding, 'ignore')

        ## ----------------
        ##  Лист изменений
        ## ----------------

        logger.out('product version definitions')
        versions = _product_version_definition(product_filename)

        v = d = ''
        for ob in versions:
            x = ob.get_version()
            if x and v != x:
                v = x
                d = getattr(ob, 'date', '')

        v = 'Product Release %s %s' % (v, d)
        _globals['PRODUCT_RELEASE'] = v.encode(default_unicode)
        logger.out('%s' % v)

        ## ----------------------------------
        ##  ГЕНЕРАЦИЯ КОНТЕНТА КОНФИГУРАТОРА
        ## ----------------------------------

        if not IsDeepDebug:
            content_generator(attrs, margins, constants, versions, params = {
                    'form'    : 'form.xml',
                    'content' : 'content.xml',
                    'script'  : 'script.txt',
                    'js'      : 'script.js',
                    'log'     : 'log.txt',
                }
            )
        if IsSyntaxControl:
            # не объявлены идентификаторы
            product_gaps = product_lexeme.difference(product_declarations)
            if len(product_gaps) > 0:
                logger.out('defs for %s product lexeme where %s are not declared' % (len(product_lexeme), len(product_gaps)))
                for x in sorted(product_gaps):
                    logger.out('%s %s' % (level1, x))
            else:
                logger.out('defs for %s product lexeme (valid)' % len(product_lexeme))
            # дублированные идентификаторы
            if duplicated['total'] > 0:
                logger.out('duplicated %s items:' % len(duplicated['ids']))
                for id in duplicated['ids']:
                    logger.out('%s %s: %s' % (level1, id, ' '.join(['%s' % str(x) for x in duplicated['items'][id]])))

        logger.close()

        if IsPrintLocale:
            locale_logger.close()

        if IsPrintCode:
            code_logger.close()

        print '--> OK'
