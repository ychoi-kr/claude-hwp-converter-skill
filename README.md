# HWP Converter

한글 문서(HWP/HWPX)를 텍스트로 변환하는 Claude Skill

## 설치 방법

1. [Releases](https://github.com/ychoi-kr/hwp-converter/releases)에서 `hwp-converter.zip` 다운로드
2. Claude 열기
3. **Settings** → **Capabilities** → **Skills**로 이동
4. **Upload Skill** 클릭하여 zip 파일 업로드

## Claude에서 사용하기

HWP/HWPX 파일을 업로드하고 다음과 같이 요청하세요:
```
이 HWP 파일을 텍스트로 변환해줘
이 문서에 뭐라고 쓰여있어?
이 HWP 파일 요약해줘
```

## 라이선스

MIT License - 자세한 내용은 [LICENSE](LICENSE) 참고

## Acknowledgements

- **Hancom Office HWP 파일 형식 명세 v5.0** - Copyright (c) Hancom Inc.
- **pyhwp** - 제어 문자 감지 알고리즘 참고 ([AGPL-3.0](https://github.com/mete0r/pyhwp))
- **olefile** - OLE 파일 형식 파싱 참고 ([BSD-2-Clause](https://github.com/decalage2/olefile))