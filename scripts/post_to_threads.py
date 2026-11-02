import os
import json
import requests

# 워크플로에서 전달받은 토큰
ACCESS_TOKEN = os.environ.get("META_ACCESS_TOKEN")

def main():
    print("🔍 원고(JSON) 파일 읽는 중...")
    with open("posts/current_post.json", "r", encoding="utf-8") as f:
        post_data = json.load(f)
    
    # 스레드용 텍스트 가져오기 (없으면 인스타 캡션으로 대체)
    threads_text = post_data.get("threads_text", post_data.get("instagram_caption", "본문 내용이 없습니다."))
    
    # 스레드 API 기본 주소
    BASE_URL = "https://graph.threads.net/v1.0"
    
    print("🔍 스레드 계정 ID 확인 중...")
    me_res = requests.get(f"{BASE_URL}/me?access_token={ACCESS_TOKEN}").json()
    if 'id' not in me_res:
        print("❌ 에러: 스레드 계정을 찾을 수 없습니다. (토큰 권한에 threads_basic, threads_content_publish가 있는지 확인하세요.)")
        print(me_res)
        return
        
    threads_user_id = me_res['id']
    
    print("📤 스레드 게시물 생성 중...")
    payload = {
        "media_type": "TEXT",
        "text": threads_text,
        "access_token": ACCESS_TOKEN
    }
    create_res = requests.post(f"{BASE_URL}/{threads_user_id}/threads", data=payload).json()
    
    if 'id' not in create_res:
        print("❌ 에러: 게시물 생성 실패")
        print(create_res)
        return
        
    creation_id = create_res['id']
    
    print("🚀 스레드에 최종 발행하는 중...")
    publish_payload = {
        "creation_id": creation_id,
        "access_token": ACCESS_TOKEN
    }
    publish_res = requests.post(f"{BASE_URL}/{threads_user_id}/threads_publish", data=publish_payload).json()
    
    if 'id' in publish_res:
        print(f"🎉 성공적으로 스레드에 자동 발행되었습니다! (게시물 ID: {publish_res['id']})")
    else:
        print(f"❌ 스레드 발행 실패: {publish_res}")

if __name__ == "__main__":
    main()
