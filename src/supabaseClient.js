import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = 'https://qrviexoisfvpxkpvqbwe.supabase.co';
// ⬇️ Supabase 콘솔 > Project Settings > API > "anon public" 키를 붙여넣기 (채팅엔 붙이지 마세요)
const SUPABASE_ANON_KEY = 'PASTE_YOUR_ANON_PUBLIC_KEY_HERE';

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
