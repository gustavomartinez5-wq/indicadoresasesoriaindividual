import { createClient } from "@supabase/supabase-js";

export const supabase = createClient(
  "https://jndvtxnzemayktytzeza.supabase.co",
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpuZHZ0eG56ZW1heWt0eXR6ZXphIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk4NDQyMzIsImV4cCI6MjA5NTQyMDIzMn0.fyiITBhwHKw2QdXIbaeJ5J_ECaPy3PcoqG12o3JT5-U"
);
