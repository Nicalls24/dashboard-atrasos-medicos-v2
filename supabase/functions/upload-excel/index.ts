import { createClient } from "npm:@supabase/supabase-js@2"
import * as XLSX from "npm:xlsx"

Deno.serve(async (req) => {
  try {
    const formData = await req.formData()

    const file = formData.get("file")

    if (!file) {
      return new Response(
        JSON.stringify({
          error: "Arquivo não enviado",
        }),
        {
          status: 400,
        }
      )
    }

    const arrayBuffer = await file.arrayBuffer()

    const workbook = XLSX.read(arrayBuffer)

    const sheet =
      workbook.Sheets[workbook.SheetNames[0]]

    const data =
      XLSX.utils.sheet_to_json(sheet)

    const supabase = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!
    )

    const chunkSize = 500

    for (
      let i = 0;
      i < data.length;
      i += chunkSize
    ) {
      const chunk =
        data.slice(i, i + chunkSize)

      const { error } = await supabase
        .from("hospital_dados")
        .upsert(chunk)

      if (error) {
        console.error(error)

        return new Response(
          JSON.stringify({
            error: error.message,
          }),
          {
            status: 500,
          }
        )
      }
    }

    return new Response(
      JSON.stringify({
        success: true,
        total: data.length,
      }),
      {
        headers: {
          "Content-Type": "application/json",
        },
      }
    )
  } catch (err) {
    return new Response(
      JSON.stringify({
        error: err.message,
      }),
      {
        status: 500,
      }
    )
  }
})