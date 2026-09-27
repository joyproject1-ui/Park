# FDA 경고장 메일 설정.
# 이 파일을 복사해 같은 폴더에 fda-mail.config.ps1 로 저장하고 값을 채우세요.
# fda-mail.config.ps1 은 .gitignore 에 있어 저장소에 올라가지 않습니다.

# 보내는 계정 (식약처 메일과 같은 계정)
$env:GMPAI_SMTP_USER     = 'joyproject1@gmail.com'

# Gmail 앱 비밀번호 16자리. 계정 비밀번호가 아닙니다.
# 식약처 스크립트가 쓰는 것을 그대로 써도 되고, myaccount.google.com/apppasswords 에서 새로 만들어도 됩니다.
$env:GMPAI_SMTP_PASSWORD = '여기에 앱 비밀번호'

# 수신자. 쉼표로 구분.
$env:GMPAI_MAIL_TO       = 'jhp5408@hanlim.com, nalkil@hanlim.com'

# 아래는 보통 손대지 않습니다.
# $env:GMPAI_SMTP_HOST   = 'smtp.gmail.com'
# $env:GMPAI_SMTP_PORT   = '587'
