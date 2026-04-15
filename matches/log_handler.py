import logging

class DatabaseLogHandler(logging.Handler):
    def emit(self, record):
        try:
            # 앱 초기화 전에 모델을 불러와 에러가 나는 것을 방지하기 위해 지연(Lazy) 임포트 적용
            from matches.models import AppLog
            
            log_entry = self.format(record)
            AppLog.objects.create(
                level=record.levelname,
                message=log_entry
            )
        except Exception:
            # 로깅 저장 중 DB가 꺼져있거나 하는 문제 발생 시, 무한 에러를 막기 위해 pass
            pass