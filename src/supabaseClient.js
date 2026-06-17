import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = 'https://qrviexoisfvpxkpvqbwe.supabase.co';
// ⬇️ Supabase 콘솔 > Project Settings > API > "anon public" 키를 붙여넣기 (채팅엔 붙이지 마세요)
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFydmlleG9pc2Z2cHhrcHZxYndlIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODE2NzA1MTAsImV4cCI6MjA5NzI0NjUxMH0.TIM9TxpS-TsjvyRbDoBUjZPOFoIXli3n6zZwjGu1UmA';

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
