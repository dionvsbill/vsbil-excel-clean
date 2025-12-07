import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

console.log('Frontend Supabase URL:', supabaseUrl);
console.log('Frontend Supabase Key:', supabaseAnonKey ? 'Loaded' : 'Missing');

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
