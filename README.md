# nexus-skills

운용 업무 자동화를 위한 Claude Code 플러그인 마켓플레이스.

## 설치 방법

이 저장소에 접근 권한이 있어야 합니다 (private). Claude Code에서:

```
/plugin marketplace add mun4046/ipo-demand-forecast-skills
```

그다음 필요한 플러그인을 설치:

```
/plugin install buycall-report@nexus-skills
```

업데이트가 있을 때는:

```
/plugin marketplace update nexus-skills
```

## 플러그인 목록

| 플러그인 | 설명 |
|---|---|
| `buycall-report` | 한국 주식 매수콜(매매전략) 보고서 생성. buy-side 관점의 sell-side 원페이저 + 상세 분석 Word(.docx) 보고서. |

## 사전 준비

- **buycall-report**: `pip install -r requirements.txt` (python-docx, pykrx, pymupdf). 스킬 폴더의 `requirements.txt` 참조. 로고·브랜드 교체는 `assets/` 및 `BUYCALL_LOGO` 환경변수 사용.
