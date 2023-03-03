import glob
import re
from datetime import datetime, timedelta, timezone

import xlrd
from ics import Event, Calendar

xlrd.Book.encoding = 'utf-8'
tz_utc_8 = timezone(timedelta(hours=8))


class QZData:
    col_start, col_end = 1, 7
    row_start, row_end = 3, 8
    new_class = []
    times_summer = [['08:00', '08:45'], ['08:55', '09:40'], ['10:10', '10:55'], ['11:05', '11:50'],
                    ['14:00', '14:45'], ['14:55', '15:40'], ['16:10', '16:55'], ['17:05', '17:50'],
                    ['19:00', '19:45'], ['19:55', '20:40'], ['20:50', '21:35']]
    times_winter = [['08:00', '08:45'], ['08:55', '09:40'], ['10:10', '10:55'], ['11:05', '11:50'],
                    ['14:30', '15:15'], ['15:25', '16:10'], ['16:40', '17:25'], ['17:35', '18:20'],
                    ['19:30', '20:15'], ['20:25', '21:10'], ['21:20', '22:05']]

    def __init__(self, week_start_data: str, file_name: str):
        self.week_start_data = datetime.strptime(re.sub('-|//', '', week_start_data), '%Y%m%d').replace(tzinfo=tz_utc_8)
        if self.week_start_data.weekday() != 0:
            print('错误：输入的日期不是星期一')
            return
        self.file_name = file_name
        self.output_filename = file_name.replace('.xls', '.ics')
        rf = xlrd.open_workbook(self.file_name)
        self.sheet = rf.sheet_by_index(0)  # 不会真的有多张表课程表混合在一起解析吧
        self.__resolve_sheet()
        self.c = Calendar()
        self.gen_new_class()
        open(self.output_filename, 'w', encoding='utf-8').write(self.c.__str__())

    def __resolve_sheet(self):
        """
        对整张表的信息解析
        :return:
        """
        for i in range(self.row_start, self.row_end):  # 去掉无效内容，直接从课表信息做解析
            for j in range(self.col_start, self.col_end):
                __i, __j = i - self.row_start, j - self.col_start  # 定义新坐标
                __class = []
                a = self.__class_resolve(__j, self.sheet.cell_value(i, j).__str__().strip())
                # print(a)
                # print('-----------------')

    def __class_resolve(self, j, class_text: str) -> list:
        if not class_text:
            return []
        __class = []
        for class_text_ in class_text.split('\n\n'):
            class_detail = ClASS(j + 1, class_text_)
            # class_detail.print()
            self.new_class.append(class_detail)
        return __class

    def gen_new_class(self):
        for class_ in self.new_class:
            class_: ClASS
            for week in class_.week:
                for time_key in class_.time_key:
                    e = Event()
                    e.name = class_.name

                    e.begin = self.week_start_data + timedelta(days=class_.week_num - 1, weeks=week - 1)
                    # print(class_.name, e.begin.month)
                    if e.begin.month in [10, 11, 12, 1, 2, 3, 4]:
                        __times = self.times_summer
                    else:
                        __times = self.times_winter
                    # __times[time_key - 1]
                    e.begin = e.begin.replace(hour=int(__times[time_key - 1][0].split(':')[0]),
                                              minute=int(__times[time_key - 1][0].split(':')[1]))
                    # print('上课时间：', __times[time_key - 1])
                    # print('上课时间：', e.begin)
                    e.end = e.begin.replace(hour=int(__times[time_key - 1][1].split(':')[0]),
                                            minute=int(__times[time_key - 1][1].split(':')[1]))
                    # print('下课时间：', e.end)

                    e.description = '教师：' + class_.teacher
                    e.location = class_.location
                    self.c.events.add(e)


class ClASS:
    def __init__(self, week_num: int, clas_text: str):
        """
        定义每节课的信息
        :param clas_text:
        """
        self.text = clas_text
        self.text_ = clas_text.split('\n')
        self.name = self.text_[0]
        self.teacher = re.sub(r'\(.*\)', '', self.text_[1])
        self.is_one_week = re.findall('单|双', self.text_[2])
        self.is_one_week = self.is_one_week[0] if self.is_one_week else '全'
        self.week_num = week_num
        self.time_key = self.__get_time_key()
        self.week = self.__get_week()
        self.location = self.text_[3]

        # print(clas_text)

    def __get_week(self):
        """
        将上课时间周范围解析为列表
        :return:
        """
        week_list = []
        week_ranges = re.sub(r'\(.*\)|\[.*]', '', self.text_[2]).replace(' ', ',')  # 获取范围数字
        for week_range in week_ranges.split(','):
            if '-' in week_range:
                start, end = week_range.split('-')
                for week in range(int(start), int(end) + 1):
                    if self.is_one_week == '单' and week % 2 == 1:
                        week_list.append(week)
                    elif self.is_one_week == '双' and week % 2 == 0:
                        week_list.append(week)
                    elif self.is_one_week == '全':
                        week_list.append(week)
            else:
                week_list.append(int(week_range))
        return week_list

    def __get_time_key(self):
        time_key = []
        week_ranges = self.text_[2].replace('(', '[').replace(')', ']')
        for time_key_ in re.search(r'[0-9][0-9](-[0-9][0-9])*节', week_ranges).group().replace('节', '').split('-'):
            time_key.append(int(time_key_))
        return time_key

    def print(self):
        # print('课程信息:', self.text)
        print('-----------------')
        print('课程名:', self.name)
        print('教师:', self.teacher)
        print('单双周:', self.is_one_week)
        print('星期:', self.week_num)
        print('节数:', self.time_key)
        print('周数:', self.week)
        print('地点:', self.location)
        # print('-----------------')
        print()


if __name__ == '__main__':
    week_start_data = '2023-02-20'
    kbs = glob.glob('*.xls')
    for kb in kbs:
        az_data = QZData(week_start_data, kb)
        print('导出', az_data.output_filename)
