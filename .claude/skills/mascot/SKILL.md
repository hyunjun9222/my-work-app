---
name: mascot
description: Generate images with imagegen.py (OpenAI image model) into images/, optionally from a reference picture. Use when the user asks for 이미지, 일러스트, 마스코트, 썸네일, 배너 for the app.
---

# mascot — AI 이미지 만들기

## 무엇을

[imagegen.py](../../../imagegen.py) 로 설명 문구를 이미지로 만든다. 결과는 `images/` 에 PNG 로 쌓인다. 앱 화면은 없애서 CLI 로만 쓴다.

## 입력

- 이미지 설명 한 문장(필수).
- 선택: `--size`(예 `1024x1536`), `--n`(장수), `--ref <참고이미지 경로>`.
- `.env` 의 `OPENAI_API_KEY`. 없으면 실행하지 말고 키가 필요하다고 알린다.

## 순서

1. 설명이 짧으면 용도·분위기·색감을 한 번 확인한다. 추측해서 살 붙이지 않는다.
2. `python imagegen.py "설명" [--size 1024x1536] [--n 2] [--ref inputs/logo.png]` 를 돌린다.
3. 저장된 파일 경로를 확인한다. 호출 비용이 있으므로 같은 요청을 반복해서 돌리지 않는다.

## 출력

만든 파일 경로 목록과 쓴 옵션을 전한다. 결과는 `images/` 폴더에서 직접 연다.

## 규칙

- 실명·실제 로고·사내 자료를 넣지 않는다.
- 사람 얼굴이 필요한 요청이면 특정 인물이 아니라 일반화된 형태로 만든다.
- 실패하면 원인(키 없음, 정책 거절 등)을 그대로 전하고 임의로 다른 그림을 만들지 않는다.
