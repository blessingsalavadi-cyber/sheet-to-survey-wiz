import { createFileRoute } from "@tanstack/react-router";
import { useState } from "react";
import { useServerFn } from "@tanstack/react-start";
import { parseExcel, type ParsedForm, type FormField } from "@/lib/excel-parse.functions";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { FileSpreadsheet, Upload, Loader2, CheckCircle2 } from "lucide-react";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Excel to Form AI" },
      { name: "description", content: "Upload an Excel file and instantly generate a Google Forms–style questionnaire from its sheets, dropdowns, and lookup tables." },
      { property: "og:title", content: "Excel to Form AI" },
      { property: "og:description", content: "Turn any .xlsx into an interactive form in seconds." },
      { property: "og:type", content: "website" },
    ],
  }),
  component: Index,
});

function Index() {
  const parseFn = useServerFn(parseExcel);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [form, setForm] = useState<ParsedForm | null>(null);
  const [answers, setAnswers] = useState<Record<string, unknown>>({});
  const [submitted, setSubmitted] = useState(false);

  async function handleFile(file: File) {
    setLoading(true);
    setError(null);
    setSubmitted(false);
    try {
      const buf = await file.arrayBuffer();
      let binary = "";
      const bytes = new Uint8Array(buf);
      const chunk = 0x8000;
      for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
      }
      const fileBase64 = btoa(binary);
      const result = await parseFn({ data: { fileBase64, fileName: file.name } });
      setForm(result);
      setAnswers({});
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to parse file");
    } finally {
      setLoading(false);
    }
  }

  if (submitted && form) {
    return (
      <div className="min-h-screen bg-background py-12 px-4">
        <div className="mx-auto max-w-2xl space-y-6">
          <Card>
            <CardHeader>
              <div className="flex items-center gap-3">
                <CheckCircle2 className="h-8 w-8 text-primary" />
                <div>
                  <CardTitle>Response recorded</CardTitle>
                  <CardDescription>Here's what you submitted.</CardDescription>
                </div>
              </div>
            </CardHeader>
            <CardContent>
              <pre className="rounded-md bg-muted p-4 text-xs overflow-auto">{JSON.stringify(answers, null, 2)}</pre>
              <div className="mt-4 flex gap-2">
                <Button onClick={() => { setSubmitted(false); setAnswers({}); }}>Fill out again</Button>
                <Button variant="outline" onClick={() => { setForm(null); setSubmitted(false); }}>Upload new file</Button>
              </div>
            </CardContent>
          </Card>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-background">
      <header className="border-b bg-card">
        <div className="mx-auto max-w-5xl px-4 py-5 flex items-center gap-3">
          <div className="flex h-9 w-9 items-center justify-center rounded-md bg-primary text-primary-foreground">
            <FileSpreadsheet className="h-5 w-5" />
          </div>
          <div>
            <h1 className="text-lg font-semibold">Excel to Form AI</h1>
            <p className="text-xs text-muted-foreground">Upload an .xlsx — get a live questionnaire</p>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-2xl px-4 py-10 space-y-6">
        {!form && (
          <Card>
            <CardHeader>
              <CardTitle>Upload an Excel file</CardTitle>
              <CardDescription>
                We'll read every sheet, detect questions, dropdowns and lookup tables, then render a form you can fill out.
              </CardDescription>
            </CardHeader>
            <CardContent>
              <label className="flex flex-col items-center justify-center rounded-lg border-2 border-dashed border-border bg-muted/40 px-6 py-12 text-center cursor-pointer hover:bg-muted transition">
                <Upload className="h-8 w-8 text-muted-foreground mb-3" />
                <span className="text-sm font-medium">Click to choose a .xlsx file</span>
                <span className="text-xs text-muted-foreground mt-1">Or drag and drop</span>
                <input
                  type="file"
                  accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  className="hidden"
                  onChange={(e) => {
                    const f = e.target.files?.[0];
                    if (f) handleFile(f);
                  }}
                />
              </label>
              {loading && (
                <div className="mt-4 flex items-center gap-2 text-sm text-muted-foreground">
                  <Loader2 className="h-4 w-4 animate-spin" /> Parsing your workbook…
                </div>
              )}
              {error && <p className="mt-4 text-sm text-destructive">{error}</p>}
            </CardContent>
          </Card>
        )}

        {form && (
          <form
            onSubmit={(e) => {
              e.preventDefault();
              setSubmitted(true);
            }}
            className="space-y-6"
          >
            <Card className="border-t-4 border-t-primary">
              <CardHeader>
                <CardTitle className="text-2xl">{form.title}</CardTitle>
                <CardDescription>Generated from {form.sections.length} sheet{form.sections.length === 1 ? "" : "s"}</CardDescription>
              </CardHeader>
            </Card>

            {form.sections.map((section, sIdx) => (
              <div key={sIdx} className="space-y-4">
                <div className="px-1">
                  <h2 className="text-sm font-semibold uppercase tracking-wide text-muted-foreground">{section.title}</h2>
                </div>
                {section.fields.map((field) => (
                  <Card key={`${sIdx}-${field.id}`}>
                    <CardContent className="pt-6 space-y-3">
                      <Label className="text-base font-medium">
                        {field.label}
                        {field.required && <span className="text-destructive ml-1">*</span>}
                      </Label>
                      <FieldInput
                        field={field}
                        value={answers[`${section.title}.${field.id}`]}
                        onChange={(v) => setAnswers((a) => ({ ...a, [`${section.title}.${field.id}`]: v }))}
                      />
                    </CardContent>
                  </Card>
                ))}
              </div>
            ))}

            <div className="flex gap-3">
              <Button type="submit">Submit</Button>
              <Button type="button" variant="outline" onClick={() => { setForm(null); setAnswers({}); }}>
                Upload different file
              </Button>
            </div>
          </form>
        )}
      </main>
    </div>
  );
}

function FieldInput({
  field,
  value,
  onChange,
}: {
  field: FormField;
  value: unknown;
  onChange: (v: unknown) => void;
}) {
  if (field.type === "select" && field.options) {
    return (
      <Select value={(value as string) ?? ""} onValueChange={onChange}>
        <SelectTrigger><SelectValue placeholder="Choose an option" /></SelectTrigger>
        <SelectContent>
          {field.options.map((opt) => (
            <SelectItem key={opt} value={opt}>{opt}</SelectItem>
          ))}
        </SelectContent>
      </Select>
    );
  }
  if (field.type === "long_text") {
    return <Textarea value={(value as string) ?? ""} onChange={(e) => onChange(e.target.value)} required={field.required} />;
  }
  if (field.type === "number") {
    return <Input type="number" value={(value as string) ?? ""} onChange={(e) => onChange(e.target.value)} required={field.required} />;
  }
  if (field.type === "date") {
    return <Input type="date" value={(value as string) ?? ""} onChange={(e) => onChange(e.target.value)} required={field.required} />;
  }
  return <Input value={(value as string) ?? ""} onChange={(e) => onChange(e.target.value)} required={field.required} />;
}
