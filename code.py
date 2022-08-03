# Два класса из page object model. Первый - для работы с такой сущностью, как проводка, второй - для работы с комбобоксами
# Невыполнение требования PEP 8 в части наименования методов/функций, а также отступ в 2 пробела - это требование работодателя

class Prov:
  """Класс для работы с проводками"""

  def __init__(self, create_date=None, operdate=None, debit=None, credit=None, amount=None, code=None, currency='KZT'):
    self.DOCNO = '--'  # (str) Номер финдока
    self.NO = '--'  # (str) Номер проводки в финдоке
    # self.DOCNO_NO = None  # (str) Номер проводки в формате DOCNO/NO
    self.CREATE_DATE = create_date  # (datetime) Дата валютирования = текущая дата (или сист дата CMS ???)
    # self._CREATE_DATE_STR = None  # ('dd/mm/yyyy hh:mm:ss') Дата валютирования
    # self.OPERDATE = None  # (date) Опердень в CMS
    self.OPERDATE_STR = operdate  # ('dd/mm/yyyy') Опердень в CMS
    self.DEBIT = debit  # (str) Счет дебета
    self.CREDIT = credit  # (str) Счет кредит
    self.AMOUNT = amount  # ('0.00') Сумма проводки
    self.CURRENCY = currency  # (KZT,USD,RUB,EUR) Валюта проводки
    self.COMMENT = None  # (str) Комментарий
    self.FULL_COMMENT = None  # (str) Полный комментарий
    #  NAME,CODE - зависимы друг от друга
    self.CODE = code  # (str) Код проводки  (PAY_CRED)
    self.NAME = Prov.code_to_name.get(code, None)  # (str) Наименование проводки  (Погашение текущего кредита)

  @property
  def DOCNO_NO(self):
    return f'{self.DOCNO}/{self.NO}'

  @property
  def CREATE_DATE_STR(self):
    return aqConvert.DateTimeToFormatStr(self.CREATE_DATE, '%d/%m/%Y %H:%M:%S')

  @property
  def OPERDATE(self):
    return aqConvert.StrToDate(self.OPERDATE_STR)

  code_to_name = {'OVER_TO_CRED': 'Перенос просроченного кредита на текущий',
                  'PAY_CRED': 'Погашение текущего кредита',
                  'MOVE_OVERDUE': 'Перенос просроченной задолженности в срочную',
                  'MOVE_OVERDUE_PRC': 'Перенос просроченных процентов в текущие',
                  'WRITE_OFF_FINECRED': 'Сторно пени за просроченный кредит',
                  'WRITE_OFF_FINEPRC': 'Сторно пени за просроченные проценты',
                  'WRITE_OFF_PEN': 'Сторно задолженности по процентам за рассрочку',
                  'WRITE_OFF': 'Списание задолженности по пене за просроченный ОД',
                  'WRITEOFF_FINE': 'Списание штрафов',
                  'WRITEOFF_PENALTY': 'Списание пени'
                  }

  def __eq__(self, other) -> bool:
    """Сравнивает две проводки"""
    allowed_time_diff = 5  # Допустимая разница во времени формирования сравниваемых проводок
    time_diff = aqDateTime.TimeInterval(self.CREATE_DATE, other.CREATE_DATE)
    time_diff = aqDateTime.GetSeconds(time_diff) + (
            aqDateTime.GetMinutes(time_diff) + aqDateTime.GetHours(time_diff) * 60) * 60
    if time_diff <= allowed_time_diff \
            and self.OPERDATE_STR == other.OPERDATE_STR \
            and self.DEBIT == other.DEBIT \
            and self.CREDIT == other.CREDIT \
            and self.AMOUNT == other.AMOUNT \
            and self.NAME == other.NAME \
            and self.CODE == other.CODE \
            and self.CURRENCY == other.CURRENCY:
      return True
    else:
      Log.Message(f'Сравнение проводок: {self} != {other}')
      return False

  def __repr__(self) -> str:
    return f'Проводка {self.DOCNO_NO} {self.CODE}({self.NAME}) Дт. {self.DEBIT} Кт. {self.CREDIT} на сумму {self.AMOUNT} {self.CURRENCY} от {self.CREATE_DATE_STR}'

  @staticmethod
  def GetProvFromFindocList(findoc_list, need_acc_id=True, result=None):
    """Метод для получения информации по проводкам из списка финдоков
    Вх.парам: need_acc_id = True (bool) - Ключ позволяющий заменить номера счетов на идентификаторы счетов;
    findoc_list (list(str)) - Список номеров финдоков
    Вых.парам: list(Prov) """
    Compass = CompassState(True)
    search_fin_doc_wnd = ItcWnd(object=Compass.StartBtn('Операции', 'Финансовые документы'))
    # Находим финдок в Compass+ и по каждой проводке создаем объект Prov
    prov_list = []
    for docno in findoc_list:
      search_fin_doc_wnd.ComboBox('Номер документа').Keys(docno)
      search_fin_doc_wnd.Button('Ввод').Click_L()
      err = Compass.PopupWindowHandler()
      if err is not None:
        CompassState.CheckExpectedResult(f'Проводка {docno} не найдена в списке финансовых документов', str(err), '',
                                         result, mode='warn_comment')
        continue
      view_findoc_wnd = ItcWnd(title='Просмотр финансового документа')
      view_findoc_table = view_findoc_wnd.Table('Дата', 2)
      operdate_str = view_findoc_wnd.ItcLineEdit('Дата:').GetText()
      for row_num in range(1, view_findoc_table.RowCount() + 1):
        if row_num != view_findoc_table.CurrentRow():
          CompassState.CheckExpectedResult(
            f'Ошибка при чтении проводок из таблицы финдока {docno}, неверный номер строки',
            view_findoc_table.CurrentRow(), row_num, result)
          break
        new_prov = Prov()
        new_prov.DOCNO = docno
        view_findoc_table.Keys('[Enter]')
        edit_fin_doc_wnd = ItcWnd(title='Корректировка проводки')
        line_edit_parent_h = edit_fin_doc_wnd.object.FindChild(['QtClassName', 'Visible', 'toolTip', 'objectName'],
                                                               ['QTCComplexEdit', True, 'Дата валютирования проводки',
                                                                'VALUEDATE_H'], 10)
        line_edit_h = str(line_edit_parent_h.FindChild('QtClassName', 'QTCLineEdit').text)
        line_edit_parent_m = edit_fin_doc_wnd.object.FindChild(['QtClassName', 'Visible', 'toolTip', 'objectName'],
                                                               ['QTCComplexEdit', True, 'Дата валютирования проводки',
                                                                'VALUEDATE_M'], 10)
        line_edit_m = str(line_edit_parent_m.FindChild('QtClassName', 'QTCLineEdit').text)
        line_edit_parent_s = edit_fin_doc_wnd.object.FindChild(['QtClassName', 'Visible', 'toolTip', 'objectName'],
                                                               ['QTCComplexEdit', True, 'Дата валютирования проводки',
                                                                'VALUEDATE_S'], 10)
        line_edit_s = str(line_edit_parent_s.FindChild('QtClassName', 'QTCLineEdit').text)
        prov_date = edit_fin_doc_wnd.ItcLineEdit('Дата валютирования:').GetText()
        prov_date_list = prov_date.split('/')
        new_prov.CREATE_DATE = aqDateTime.SetDateTimeElements(int(prov_date_list[2]), int(prov_date_list[1]),
                                                              int(prov_date_list[0]), int(line_edit_h),
                                                              int(line_edit_m), int(line_edit_s))
        new_prov.OPERDATE_STR = operdate_str
        new_prov.DEBIT = edit_fin_doc_wnd.ItcLineEdit('Дебет:').GetText()
        new_prov.CREDIT = edit_fin_doc_wnd.ItcLineEdit('Кредит:').GetText()
        new_prov.AMOUNT = CompassState.Float2Str(edit_fin_doc_wnd.ItcLineEdit('Сумма:').GetText())
        new_prov.COMMENT = edit_fin_doc_wnd.ItcLineEdit('Комментарий:').GetText()
        new_prov.CODE = edit_fin_doc_wnd.ComboBox('Код проводки:').GetCurrentText()
        edit_fin_doc_wnd.ItcWndClose()
        current_row = view_findoc_table.GetRow()
        new_prov.NO = current_row[0]
        new_prov.CURRENCY = CompassState.CodeToCurrency(current_row[4])
        new_prov.FULL_COMMENT = view_findoc_wnd.ItcLineEdit('Полн. коммент.:').GetText()
        new_prov.NAME = view_findoc_wnd.ItcLineEdit('Проводка:').GetText()
        prov_list.append(new_prov)
        view_findoc_table.Keys('[Down]')
      view_findoc_wnd.ItcWndClose()
    search_fin_doc_wnd.ItcWndClose()
    if need_acc_id:
      # Словарь для хранения найденных идентификаторов счетов. Технические счета отсавляем как есть
      all_acc_id = {k: v for k, v in zip(CompassState.TECH_ACC, CompassState.TECH_ACC)}
      search_acc_wnd = ItcWnd(object=Compass.StartBtn('Операции', 'Счета физических лиц'))
      for prov in prov_list:
        for acc in (prov.DEBIT, prov.CREDIT):
          if acc not in all_acc_id:
            search_acc_wnd.ComboBox('Номер счета:').Keys(acc)
            search_acc_wnd.Button('Ввод').Click_L()
            if Compass.PopupWindowHandler() is not None:
              all_acc_id[acc] = acc
            else:
              acc_wnd = ItcWnd(title='Счет ')  # Пробел обязателен, чтобы исключить окно "Счета физ лиц"
              acc_wnd.MunuBar_CallItem('Данные', 'Атрибуты счета')
              # Редкая ошибка возникающая при открытии атрибутов счета с формы поиска счетов
              if Compass.PopupWindowHandler() is not None:
                acc_wnd.ItcWndClose()
                continue
              atribut_wnd = ItcWnd(title='Атрибуты счета')
              atribut_wnd.ItcPageList_OpenTab('Договор')
              dea_num = atribut_wnd.ItcLineEdit('Номер:').GetText()
              atribut_wnd.ItcWndClose()
              acc_wnd.ItcWndClose()
              dea_wnd = Compass.SearchAndOpenClientsDea(dea_num)
              dea_wnd.ItcPageList_OpenTab('Счета')
              acc_table = dea_wnd.Table('Счет').GetLastRow(0)
              dea_wnd.ItcPageList_OpenTab('Связанные договоры')
              dea_num2 = dea_wnd.Table('Номер договора').GetRow()[0]
              dea_wnd.Table('Номер договора').OpenRow('LAST_ROW')
              dea_wnd2 = ItcWnd(full_title=f'Клиентский договор [{dea_num2}]')
              dea_wnd2.ItcPageList_OpenTab('Счета')
              Delay(3000)  # TODO: Редкая ошибка, не находит dea_wnd2.Table('Счет')
              acc_table.extend(dea_wnd2.Table('Счет').GetLastRow(0))
              dea_wnd2.ItcWndClose()
              dea_wnd.ItcWndClose()
              ItcWnd(title='Договоры физических лиц').ItcWndClose()
              for acc_row in acc_table:
                all_acc_id[acc_row['Счет']] = acc_row['Идентификатор']
      for prov in prov_list:
        prov.DEBIT = all_acc_id.get(prov.DEBIT, prov.DEBIT)
        prov.CREDIT = all_acc_id.get(prov.CREDIT, prov.CREDIT)
      search_acc_wnd.ItcWndClose()
    Log.Event(f'Сформировано {len(prov_list)} проводок >>', str(prov_list))
    return prov_list

  @staticmethod
  def CheckProvInFindoc(**kwargs):
    """
    Проверка наличия проводки в финансовом документе
    Вх.парам: kwargs['Финдок'] (str) - Номер финдока, в котором будет произведен поиск
    Необязательные параметры, по которым будет произведен поиск проводки. Любой атрибут объекта Prov. Если параметр == '', он будет проигнорирован
    kwargs['DEBIT']
    kwargs['CREDIT']
    kwargs['AMOUNT']
    kwargs['CURRENCY']
    kwargs['COMMENT']
    kwargs['FULL_COMMENT']
    kwargs['CODE']
    kwargs['NAME']
    Вых.парам: {'Результат операции': 'OK или текст ошибки', 'Искомая проводка': str, 'Список проводок': str(list), 'Количество проводок': int}
    """
    result = {}
    provs = Prov.GetProvFromFindocList([kwargs.pop('Финдок', '')], True)
    result['Искомая проводка'] = None
    result['Список проводок'] = str(provs)
    result['Количество проводок'] = len(provs)
    for prov in provs:
      for key in kwargs:
        if key == 'AMOUNT':
          kwargs['AMOUNT'] = CompassState.Float2Str(kwargs['AMOUNT'])
        if kwargs[key] != '' and kwargs[key] != getattr(prov, key, 'Несуществующий атрибут'):
          break
      else:
        result['Результат операции'] = 'OK'
        result['Искомая проводка'] = str(prov)
        break
    else:
      result['Результат операции'] = 'Проводка не найдена'
    return result
    
    
    
    
    
    
    
class ComboBoxClass(ElementClass):
  """ Класс для комбобоксов """

  def __init__(self, parent_wnd, combo_box_object, combo_box_name):
    super().__init__(parent_wnd, combo_box_object)
    self.el_type = 'Комбобокс'
    self.el_name = combo_box_name

  def Select_Item(self, item_name):
    """ Устанавливает комбобокс в требуемое значение
    Вх.параметры: item_name (str) - требуемое значение комбобокса (номер из первого столбца)
    Вых.параметры: True/False
    """
    self.item_name = item_name
    if self.element.Exists:
      if str(self.element.currentText) != self.item_name:
        self.element.QComboBox_showPopup()  # Вызываем выпадающий список
        self.list_box = Sys.Process("itc5").QtObject("QComboBoxPrivateContainer", "", 1).UIAObject("QTCComboListBox")
        if self.list_box.VisibleOnScreen:
          self.Keys('[Down]')  # Активируем выпадающий список
          self.Keys('[Home]')
          for _ in range(0, self.list_box.ChildCount):  # Ищем необходимый пункт в выпадающем списке
            list_box_item = self.list_box.Findchild('IsSelected', True)
            if list_box_item.Exists and list_box_item.NativeUIAObject.Name[:len(self.item_name)] == self.item_name:
              self.Keys('[Enter]')
              break
            self.Keys('[Down]')
      if str(self.element.currentText) == self.item_name:
        Log.Event('Комбобокс "' + self.el_name + '", значение "' + self.item_name + '" установлено')
        return True
      else:
        Log.Warning('Комбобокс "' + self.el_name + '", значение "' + self.item_name + '" не найдено')
    else:
      Log.Warning('Комбобокс "' + self.el_name + '" не найден')
    return False
    
