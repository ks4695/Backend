const axios = require('axios');
const fs = require('fs');

const CLIENT_ID = '6e3b85ad-a49b-4259-ae1e-bcb870e41474';
const TOKEN_FILE = './onedrive_tokens.json';

async function getToken() {
  console.log('\n📱 Device Code Flow 시작...\n');

  // 1. Device Code 요청
  const deviceCodeResponse = await axios.post(
    'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode',
    new URLSearchParams({
      client_id: CLIENT_ID,
      scope: 'Files.ReadWrite Files.ReadWrite.All Mail.Send offline_access'
    }),
    { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
  );

  const { user_code, device_code, verification_uri, expires_in, interval } = deviceCodeResponse.data;

  console.log('====================================================');
  console.log('🔐 아래 단계를 따라하세요:');
  console.log('====================================================');
  console.log(`\n1. 브라우저에서 이 URL 접속:`);
  console.log(`   👉 ${verification_uri}`);
  console.log(`\n2. 이 코드를 입력하세요:`);
  console.log(`   👉 ${user_code}`);
  console.log(`\n3. OneDrive Microsoft 계정으로 로그인`);
  console.log(`\n⏰ ${expires_in}초 안에 완료하세요\n`);
  console.log('로그인 대기 중');

  // 2. 폴링 (로그인 완료될 때까지 기다림)
  const pollInterval = (interval || 5) * 1000;
  const maxAttempts = Math.floor(expires_in / (interval || 5));

  for (let i = 0; i < maxAttempts; i++) {
    await new Promise(resolve => setTimeout(resolve, pollInterval));
    process.stdout.write('.');

    try {
      const tokenResponse = await axios.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        new URLSearchParams({
          client_id: CLIENT_ID,
          grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
          device_code: device_code
        }),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
      );

      const tokens = {
        access_token: tokenResponse.data.access_token,
        refresh_token: tokenResponse.data.refresh_token,
        expires_at: Date.now() + (tokenResponse.data.expires_in * 1000)
      };

      fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokens, null, 2));

      console.log('\n\n✅ 로그인 성공!');
      console.log('====================================================');
      console.log('🔑 Render에 넣을 REFRESH_TOKEN:');
      console.log('====================================================');
      console.log(tokens.refresh_token);
      console.log('====================================================');
      console.log('\n👆 위 값을 복사해서 Render 환경변수 REFRESH_TOKEN에 넣으세요!\n');
      return;

    } catch (err) {
      if (err.response?.data?.error === 'authorization_pending') {
        // 아직 로그인 안 함, 계속 대기
      } else if (err.response?.data?.error === 'authorization_declined') {
        console.log('\n❌ 로그인 거부됨');
        return;
      } else {
        console.error('\n❌ 오류:', err.response?.data || err.message);
        return;
      }
    }
  }

  console.log('\n❌ 시간 초과');
}

getToken().catch(console.error);
