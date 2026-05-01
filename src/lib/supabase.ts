import { createClient } from "@supabase/supabase-js";

const supabaseUrl = (import.meta.env.VITE_SUPABASE_URL ?? "").trim();
const supabaseAnonKey = (import.meta.env.VITE_SUPABASE_ANON_KEY ?? "").trim();

export const supabaseConfigError =
  !supabaseUrl || !supabaseAnonKey
    ? "请先在项目根目录 .env.local 填写 VITE_SUPABASE_URL 和 VITE_SUPABASE_ANON_KEY"
    : null;

export const supabase = supabaseConfigError ? null : createClient(supabaseUrl, supabaseAnonKey);
