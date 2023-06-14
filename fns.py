import zeep
import base64
import time
from lxml import etree
import datetime
import data_worker as dw

class FNSApi:

    def __init__(self):
        self.master_token = dw.get_parameter('master_token')
        self.app_id = dw.get_parameter('app_id')
        self.point = dw.get_parameter('point')
        date_counter = dw.get_parameter('date_counter')
        now = str(datetime.datetime.today().date())
        if date_counter == now:
            self.counter = int(dw.get_parameter('counter'))
        else:
            self.counter = 0
            dw.save_parameter('date_counter', now)


    def get_session_token(self):
        wsdl = self.point + 'open-api/AuthService/0.1?wsdl'
        xml = '<tns:AuthRequest xmlns:tns="urn://x-artefacts-gnivc-ru/ais3/kkt/AuthService/types/1.0"><tns:AuthAppInfo><tns:MasterToken>' + \
              self.master_token + '</tns:MasterToken></tns:AuthAppInfo></tns:AuthRequest>'
        root = etree.fromstring(xml)
        client = zeep.Client(wsdl=wsdl)
        response = client.service.GetMessage(Message={'_value_1': root})
        self.counter += 1
        tag = response[0].tag
        result = {
            'status': 'error',
            'message': '',
            'token': None,
            'expire_time': None
        }
        if tag == '{urn://x-artefacts-gnivc-ru/ais3/kkt/AuthService/types/1.0}Fault':
            result['message'] = response[0][0].text
        elif tag == '{urn://x-artefacts-gnivc-ru/ais3/kkt/AuthService/types/1.0}Result':
            self.session_token = response[0][0].text
            result['status'] = 'success'
            result['message'] = 'Токен выдан'
            result['expire_time'] = response[0][1].text
        return result



    def process_ticket(self, request_type, sum, timestamp, fiscal_number, operation_type, fiscal_document_id, fiscal_sign):
        result = {
            'status': 'error',
            'message': '',
            'code': None,
        }
        wsdl = self.point + 'open-api/ais3/KktService/0.1?wsdl'
        xml = """
        <tns:{request_type}TicketRequest xmlns:tns="urn://x-artefacts-gnivc-ru/ais3/kkt/KktTicketService/types/1.0">
            <tns:{request_type}TicketInfo>
                <tns:Sum>{sum}</tns:Sum>
                <tns:Date>{timestamp}</tns:Date>
                <tns:Fn>{fiscal_number}</tns:Fn>
                <tns:TypeOperation>{operation_type}</tns:TypeOperation>
                <tns:FiscalDocumentId>{fiscal_document_id}</tns:FiscalDocumentId>
                <tns:FiscalSign>{fiscal_sign}</tns:FiscalSign>
            </tns:{request_type}TicketInfo>
        </tns:{request_type}TicketRequest>
        """.format(
            request_type=request_type,
            sum=sum.strip(),
            timestamp=timestamp.strip(),
            fiscal_number=fiscal_number.strip(),
            operation_type=operation_type.strip(),
            fiscal_document_id=fiscal_document_id.strip(),
            fiscal_sign=fiscal_sign.strip()
        )
        root = etree.fromstring(xml)
        client = zeep.Client(wsdl=wsdl)


        client.transport.session.headers.update({
            'FNS-OpenApi-Token': self.session_token,
            'FNS-OpenApi-UserToken': base64.b64encode(self.app_id.encode('ascii'))
        })
        messageId = client.service.SendMessage(Message={'_value_1': root})
        wait_count = 0
        while True:
            time.sleep(2)
            wait_count += 2
            response = client.service.GetMessage(MessageId=messageId)
            self.counter += 1
            if (response['ProcessingStatus'] == 'COMPLETED'):
                break
            if wait_count > 40:
                result['message'] = "Превышено время ожидания ответа от ФНС "
                return result

        response = response['Message']['_value_1']
        tag = response[0].tag
        if tag == '{urn://x-artefacts-gnivc-ru/ais3/kkt/KktTicketService/types/1.0}Fault':
            result['message'] = response[0][0].text
        elif tag == '{urn://x-artefacts-gnivc-ru/ais3/kkt/KktTicketService/types/1.0}Result':
            result['status'] = 'success'
            result['code'] = response[0][0].text
            result['message'] = response[0][1].text
        time.sleep(1)
        return result



    def check_ticket(self,sum, timestamp, fiscal_number, operation_type, fiscal_document_id, fiscal_sign):
        return self.process_ticket('Check',sum, timestamp, fiscal_number, operation_type, fiscal_document_id, fiscal_sign)

    def get_ticket(self,sum, timestamp, fiscal_number, operation_type, fiscal_document_id, fiscal_sign):
        return self.process_ticket('Get', sum, timestamp, fiscal_number, operation_type, fiscal_document_id, fiscal_sign)

    def set_counter(self):
        dw.save_parameter('counter', self.counter)