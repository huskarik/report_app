from requests import Session
from datetime import datetime
import time
from app.excel_filler import fill_excel_report
from app.logger import setup_logger
import os

logger = setup_logger(__name__)

# --- Жёсткое сопоставление проекта с контрагентом --- (ключ проект, значение - контрагент)
MAP_PROJECT_AGENT = {
    "https://api.moysklad.ru/api/remap/1.2/entity/project/9eecc057-dc57-11ee-0a80-1406000914a9":
        "https://api.moysklad.ru/api/remap/1.2/entity/counterparty/9e082119-cecc-11f0-0a80-177e00540242", # OZON
    "https://api.moysklad.ru/api/remap/1.2/entity/project/b0fc98ef-ec4c-11ee-0a80-17510017c31f":
        "https://api.moysklad.ru/api/remap/1.2/entity/counterparty/fa059598-cecc-11f0-0a80-11e900544c25", # WB
    "https://api.moysklad.ru/api/remap/1.2/entity/project/a2fce344-bee3-11f0-0a80-0311000b7b12":
        "https://api.moysklad.ru/api/remap/1.2/entity/counterparty/c15d626b-1189-11f1-0a80-0338004494a9" # YANDEX
}


class ReportGenerator:
    def __init__(self, token):
        self.token = token
        self.base_url_demand = "https://api.moysklad.ru/api/remap/1.2/entity/demand"
        self.base_url_comission_report = "https://api.moysklad.ru/api/remap/1.2/entity/comission_report"
        self.base_url_refound = "https://api.moysklad.ru/api/remap/1.2/entity/salesreturn"

        self.session = Session()

        self.url_filtered_demands = ""
        self.url_comission_report = ""
        self.url_filtered_refounds = ""

        self.headers = {
            "Authorization": f"Bearer {self.token}",
        }

        self.current_demand_numbers = []
        self.current_positions_in_demands = []

        self.current_refound_numbers = []
        self.current_positions_in_refounds = []

        self.current_comission_numbers = []
        self.current_positions_in_comission = []
        self.current_refounds_in_comission = []

        self.curr_from_date = ""
        self.curr_to_date = ""

    def set_urls(self, project=None, from_date=None, to_date=None):
        agent_url = MAP_PROJECT_AGENT.get(project)
        # Устанавливаем урл для всех получения всех отгрузок
        self.url_filtered_demands = self.base_url_demand + "?filter=project=" + project + \
        ";moment>=" + from_date + ";moment<=" + to_date + "&order=name,desc"

        # Устанавливаем урл для получения отчётов
        self.url_comission_report = \
            f"https://api.moysklad.ru/api/remap/1.2/entity/commissionreportin?filter=agent={agent_url}&order=name,desc"

        # Устанавливаем урл для получения возвратов покупателей
        self.url_filtered_refounds = self.base_url_refound + "?filter=project=" + project + \
        ";moment>=" + from_date + ";moment<=" + to_date + "&order=name,desc"

        # Устанавливаем даты
        self.curr_from_date = datetime.fromisoformat(from_date)
        self.curr_to_date = datetime.fromisoformat(to_date)

    def __make_request(self, method, url, headers, body=None):
        r = self.session.request(method=method, url=url, headers=headers, data=body)

        return r.json()

    def get_demands(self):
        # Запрос в мой склад, получить все
        all_demands = self.__make_request(method="GET", url=self.url_filtered_demands, headers=self.headers)

        # Идём по каждой отгрузке
        for row in all_demands.get("rows"):
            positions_url = row.get("positions").get("meta").get("href") if row.get("positions") else None

            # Если позиции отсутствуют, переходим к следующему документу
            if positions_url is None:
                continue

            self.current_demand_numbers.append("Отгрузка № " + row.get("name"))

            # Запрашиваем позиции документа
            demands_positions = self.__make_request(method="GET", url=positions_url + "?expand=assortment", headers=self.headers)

            self.__fill_local_positions(
                self.current_positions_in_demands,
                demands_positions.get("rows")
            )

        logger.info("DEMANDS TRUE")

    def get_comission_reports(self):
        all_comission_reports = self.__make_request(method="GET", url=self.url_comission_report, headers=self.headers)

        target_report_list = list()

        # Выбираем отчёты комиссионера подходящие под период
        if all_comission_reports.get("rows"):
            for report in all_comission_reports.get("rows"):
                start = report.get("commissionPeriodStart")
                end = report.get("commissionPeriodEnd")

                if start:
                    start = datetime.fromisoformat(start.replace(" ", "T"))

                if end:
                    end = datetime.fromisoformat(end.replace(" ", "T"))

                # Захватываем периоды в отчёте комиссионера
                if start and end:
                    if end >= self.curr_from_date and start <= self.curr_to_date:
                        target_report_list.append(report)

        if target_report_list:
            for report in target_report_list:

                # Инициализируем переменные
                positions = dict()
                refounds = dict()

                # Получили проданные позиции из отчёта комиссионера
                url_positions = report.get('positions').get("meta").get("href")
                if url_positions:
                    positions = self.__make_request(
                        method="GET",
                        url=url_positions + "?expand=assortment",
                        headers=self.headers,
                    )
                # Получили возвращённые позиции из отчёта комиссионера
                url_refounds = report.get('returnToCommissionerPositions').get("meta").get("href")
                if url_refounds:
                    refounds = self.__make_request(
                        method="GET",
                        url=url_refounds + "?expand=assortment",
                        headers=self.headers
                    )
                # Записываем номера документов
                if refounds.get("rows") or positions.get("rows"):
                    self.current_comission_numbers.append("Отчёт комиссионера № " + report.get("name"))
                else:
                    continue

                # Сначала идём по проданным позициям
                self.__fill_local_positions(
                    self.current_positions_in_comission,
                    positions.get("rows")
                )

                # Потом по возвратам в отчёте комиссионера
                self.__fill_local_positions(
                    self.current_refounds_in_comission,
                    refounds.get("rows"),
                    is_refound=True
                )

        logger.info("COMMISSION TRUE")

    def get_refounds(self):
        # Запрос в мой склад, получить все
        all_demands = self.__make_request(
            method="GET",
            url=self.url_filtered_refounds,
            headers=self.headers)

        # Идём по каждой отгрузке
        for row in all_demands.get("rows"):
            positions_url = row.get("positions").get("meta").get("href") if row.get("positions") else None

            # Если позиции отсутствуют, переходим к следующему документу
            if positions_url is None:
                continue

            self.current_refound_numbers.append("Возврат покупателя № " + row.get("name"))

            # Запрашиваем позиции документа
            refounds_positions = self.__make_request(
                method="GET",
                url=positions_url + "?expand=assortment",
                headers=self.headers)

            self.__fill_local_positions(
                self.current_positions_in_refounds,
                refounds_positions.get("rows"),
                is_refound=True,
                nsp=True
            )

        logger.info("REFOUNDS TRUE")

    def __fill_local_positions(self, container_positions, rows, is_refound=False, nsp=False):
        for row in rows:
            temp_position = dict()

            if is_refound:
                temp_position["art"] = row.get("assortment").get("article")
                temp_position["name"] = row.get("assortment").get("name")
                temp_position["price"] = -float(row.get("price")) / 100
                temp_position["quantity"] = -float(row.get("quantity"))
                if nsp:
                    temp_position["NSP"] = "НСП"
            else:
                temp_position["art"] = row.get("assortment").get("article")
                temp_position["name"] = row.get("assortment").get("name")
                temp_position["price"] = float(row.get("price")) / 100
                temp_position["quantity"] = float(row.get("quantity"))

            container_positions.append(temp_position)

    def generate_report(self, project: str):
        self.get_demands()
        self.get_comission_reports()
        self.get_refounds()

        logger.info("##############################################")
        logger.info(f"DOC NUMBERS DEMANDS  {self.current_demand_numbers}")
        logger.info(f"COUNT DOCS DEMANDS {len(self.current_demand_numbers)}")
        logger.info(f"TOTAL DEMANDS POSITIONS {len(self.current_positions_in_demands)} \n\n")

        logger.info(f"DOC NUMBERS COMISSION {self.current_comission_numbers}")
        logger.info(f"COUNT DOCS COMISSIONS {len(self.current_comission_numbers)}")
        logger.info(f"TOTAL COMISSION POSITIONS {len(self.current_positions_in_comission)}")
        logger.info(f"TOTAL COMISSION REFOUNDS POSITIONS {len(self.current_refounds_in_comission)} \n\n")

        logger.info(f"DOC NUMBERS REFOUNDS {self.current_refound_numbers}")
        logger.info(f"COUNT DOCS REFOUNDS {len(self.current_refound_numbers)}")
        logger.info(f"TOTAL REFOUNDS POSITIONS {len(self.current_positions_in_refounds)} \n\n")

        #max_height = max(len(self.current_positions_in_demands), len(self.current_refound_numbers))

        sections = {
            'A': {'start_row': 5, 'data': self.current_positions_in_demands},  # колонки 1-4
            'B': {
                'start_row': 5,
                'data': self.current_positions_in_comission +
                        self.current_refounds_in_comission +
                        self.current_positions_in_refounds
            },  # колонки 5-8
            'C': {'start_row': 8, 'data': self.current_demand_numbers},
            'D': {'start_row': 10, 'data': self.current_comission_numbers + self.current_refound_numbers}
        }

        timestamp = int(time.time())
        filename = f"report_{timestamp}.xlsx"

        # --- Создаём папку temp, если её нет ---
        base_dir = os.path.dirname(os.path.abspath(__file__))  # это app/
        temp_dir = os.path.join(base_dir, "temp")
        os.makedirs(temp_dir, exist_ok=True)

        # Полный путь до отчёта
        filepath = os.path.join(temp_dir, filename)

        fill_excel_report(
            template_path=os.path.join(base_dir, 'шаблон.xlsx'),
            output_path=filepath,
            sections=sections,
            project_name=project,
            from_date=self.curr_from_date,
            to_date=self.curr_to_date
        )
        return filename










