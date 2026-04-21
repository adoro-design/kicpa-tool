============================================================
KICPA 콘텐츠 관리 툴 - 설치 안내 (GitHub 배포 방식)
============================================================

■ 사전 준비
  - GitHub 계정
  - 서버 SSH 접속 가능
  - 서버에 PHP 8.0 이상, Composer, Git 설치되어 있어야 함

■ 설치 순서

  [1단계] GitHub에 저장소 생성 (비공개)
    - GitHub Desktop → File → New Repository
    - 이름: kicpa-tool / Private 선택 후 생성

  [2단계] 코드 Push
    - GitHub Desktop에서 이 kicpa-tool 폴더를 저장소로 추가
    - Commit → Push origin

  [3단계] 서버에서 Clone
    SSH 접속 후:
    $ cd /public_html
    $ git clone https://github.com/[계정명]/kicpa-tool.git
    $ cd kicpa-tool
    $ composer install --no-dev

  [4단계] config.php 설정
    $ cp config.example.php config.php
    $ nano config.php

    수정 항목:
      DB_HOST = 'localhost'
      DB_NAME = '데이터베이스명'
      DB_USER = 'DB사용자명'
      DB_PASS = 'DB비밀번호'
      BASE_URL = '/kicpa-tool'

  [5단계] DB 초기화
    브라우저: http://도메인/kicpa-tool/install.php

  [6단계] install.php 삭제
    $ rm install.php

■ 이후 업데이트
    [PC] GitHub Desktop으로 수정사항 Push
    [서버] $ cd /public_html/kicpa-tool && git pull

■ 기본 로그인
    아이디: admin  /  비밀번호: kicpa1234\!
    ※ 로그인 후 즉시 비밀번호 변경 필수\!

============================================================
