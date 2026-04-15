from django.db import models
from django.contrib.auth.models import User

# 1. 회원 프로필 (조, 성별 정보 포함)
class Profile(models.Model):
    GROUP_CHOICES = [('A', 'A조(10년)'), ('B', 'B조(5년)'), ('C', 'C조(3년미만)')]
    GENDER_CHOICES = [('M', '남성'), ('F', '여성')]
    
    # 🌟 게스트는 User가 없으므로 null=True, blank=True 추가
    user = models.OneToOneField(User, on_delete=models.CASCADE, null=True, blank=True)
    is_guest = models.BooleanField(default=False, verbose_name="게스트 여부")
    
    name = models.CharField(max_length=20)
    group = models.CharField(max_length=1, choices=GROUP_CHOICES)
    gender = models.CharField(max_length=1, choices=GENDER_CHOICES)
    is_owner = models.BooleanField(default=False, verbose_name="사장님 권한")

    def __str__(self):
        # 🌟 이름 옆에 게스트 표시 추가
        status = "[게스트]" if self.is_guest else ""
        return f"{self.name}{status} ({self.group})"

# 2. 월례회 모임 날짜 및 마감 상태
class MonthlyMeet(models.Model):
    # 기존 필드
    date = models.DateField(unique=True) 
    title = models.CharField(max_length=100)
    
    # 신규 추가 필드: 오늘 모임의 마감 여부를 결정합니다.
    is_finalized = models.BooleanField(default=False, verbose_name="마감 여부")

    def __str__(self):
        return self.title

# 3. 경기 기록
class Match(models.Model):
    meet = models.ForeignKey(MonthlyMeet, on_delete=models.CASCADE, related_name='matches')
    court = models.IntegerField(choices=[(2, '2코트'), (3, '3코트')])
    
    player1 = models.ForeignKey(Profile, on_delete=models.CASCADE, related_name='matches_as_p1')
    player2 = models.ForeignKey(Profile, on_delete=models.CASCADE, related_name='matches_as_p2')
    
    p1_score = models.IntegerField(null=True, blank=True)
    p2_score = models.IntegerField(null=True, blank=True)
    
    applied_handicap = models.IntegerField(default=0) # 적용된 핸디캡 점수
    is_completed = models.BooleanField(default=False) # 경기 완료 여부
    
    # 🌟 신규 추가: 누가 점수를 처음 입력했는지 추적하는 필드
    recorded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="점수 입력자")

# 4. 공지사항 (Notice) 모델 (들여쓰기 수정됨)
class Notice(models.Model):
    title = models.CharField(max_length=200)
    content = models.TextField()
    created_at = models.DateTimeField(auto_now_add=True)
    is_important = models.BooleanField(default=False) # 중요 공지 여부 (상단 고정용)
    view_count = models.PositiveIntegerField(default=0, verbose_name="조회수")

    # 🌟 [신규] 작성자, 작성자명 재정의, 위치 정보 필드 추가
    author = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="작성자")
    author_display_name = models.CharField(max_length=50, blank=True, null=True, verbose_name="작성자명(관리자용)")
    location_name = models.CharField(max_length=100, blank=True, null=True, verbose_name="장소/위치 이름")

    def __str__(self):
        return self.title

    # 🌟 [신규] 화면에 표시될 최종 작성자 이름을 결정하는 함수
    def get_author_name(self):
        if self.author_display_name:
            return self.author_display_name
        if self.author:
            if hasattr(self.author, 'profile') and self.author.profile.is_owner:
                return "사장님"
            if hasattr(self.author, 'profile'):
                return self.author.profile.name
            return self.author.username
        return "알 수 없음"

# 🌟 [신규] 공지사항 댓글 모델
class NoticeComment(models.Model):
    notice = models.ForeignKey(Notice, on_delete=models.CASCADE, related_name='comments')
    author = models.ForeignKey(User, on_delete=models.CASCADE, verbose_name="댓글 작성자")
    content = models.TextField(verbose_name="댓글 내용")
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        if self.author.profile:
            return f"'{self.notice.title}'의 댓글 by {self.author.profile.name}"
        return f"'{self.notice.title}'의 댓글 by {self.author.username}"

# 🌟 [신규] 시스템/에러 로그를 DB에 저장하기 위한 모델
class AppLog(models.Model):
    level = models.CharField(max_length=20, verbose_name="로그 레벨")
    message = models.TextField(verbose_name="로그 내용")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="발생 일시")

    class Meta:
        verbose_name = "시스템 로그"
        verbose_name_plural = "시스템 로그 목록"

    def __str__(self):
        return f"[{self.level}] {self.message[:50]}"