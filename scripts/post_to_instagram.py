import os
import json
import requests
import time

# GitHub Secrets에서 가져온 토큰과 저장소 정보
ACCESS_TOKEN = os.environ.get("IG_ACCESS_TOKEN")
GITHUB_REPO = os.environ.get("GITHUB_REPOSITORY")  # 예: username/brunch-sns-auto
BRANCH = "main"

# JSON 파일에서 인스타 캡션(본문) 읽어오기
with open("posts/current_post.json", "r", encoding="utf-8") as f:
    post_data = json.load(f)
caption = post_data.get("instagram_caption", "본문 내용이 없습니다.")

GRAPH_URL = "https://graph.facebook.com/v18.0"

def get_ig_user_id():
    # 1. 페이스북 페이지 ID 가져오기
    pages_res = requests.get(f"{GRAPH_URL}/me/accounts?access_token={ACCESS_TOKEN}").json()
    if 'data' not in pages_res or len(pages_res['data']) == 0:
        raise Exception("연결된 페이스북 페이지를 찾을 수 없습니다. (토큰 권한 확인 필요)")
    page_id = pages_res['data'][0]['id']
    
    # 2. 인스타그램 비즈니스 계정 ID 가져오기
    ig_res = requests.get(f"{GRAPH_URL}/{page_id}?fields=instagram_business_account&access_token={ACCESS_TOKEN}").json()
    if 'instagram_business_account' not in ig_res:
        raise Exception("페이지에 연결된 인스타그램 비즈니스 계정을 찾을 수 없습니다.")
    return ig_res['instagram_business_account']['id']

def main():
    print("🔍 인스타그램 계정 ID를 찾는 중...")
    ig_user_id = get_ig_user_id()
    
    print("📤 카드뉴스 5장을 인스타그램 서버에 업로드하는 중 (시간이 조금 걸립니다)...")
    container_ids = []
    
    # 1~5번 슬라이드를 각각 개별 업로드 (Carousel Item)
    for i in range(1, 6):
        # GitHub에 저장된 이미지의 Raw URL 생성 (저장소가 Public이어야 함)
        img_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/{BRANCH}/posts/images/slide-{i}.png"
        payload = {
            "image_url": img_url,
            "is_carousel_item": "true",
            "access_token": ACCESS_TOKEN
        }
        res = requests.post(f"{GRAPH_URL}/{ig_user_id}/media", data=payload).json()
        if 'id' in res:
            container_ids.append(res['id'])
            print(f" - {i}번 슬라이드 업로드 성공")
        else:
            print(f" - {i}번 슬라이드 에러:", res)

    # 인스타그램 서버가 이미지를 처리할 시간을 줌
    time.sleep(10)

    print("🔗 슬라이드를 하나의 게시물(Carousel)로 묶는 중...")
    carousel_payload = {
        "media_type": "CAROUSEL",
        "children": ",".join(container_ids),
        "caption": caption,
        "access_token": ACCESS_TOKEN
    }
    carousel_res = requests.post(f"{GRAPH_URL}/{ig_user_id}/media", data=carousel_payload).json()
    creation_id = carousel_res['id']

    print("🚀 최종 게시물을 인스타그램에 발행하는 중...")
    publish_payload = {
        "creation_id": creation_id,
        "access_token": ACCESS_TOKEN
    }
    publish_res = requests.post(f"{GRAPH_URL}/{ig_user_id}/media_publish", data=publish_payload).json()
    
    print(f"🎉 성공적으로 인스타그램에 자동 발행되었습니다! (게시물 ID: {publish_res.get('id')})")

if __name__ == "__main__":
    main()
