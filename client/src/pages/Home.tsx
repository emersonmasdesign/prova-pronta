import { useEffect, useMemo, useRef, useState } from "react";
import {
  AlignJustify,
  ArrowDown,
  ArrowUp,
  Check,
  ChevronDown,
  CircleHelp,
  Download,
  FileDown,
  FileText,
  ImagePlus,
  Info,
  LayoutGrid,
  Minus,
  Plus,
  Printer,
  RotateCcw,
  Save,
  Settings2,
  ShieldCheck,
  KeyRound,
  Shuffle,
  Sparkles,
  Trash2,
  Upload,
  X,
} from "lucide-react";
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import { toast } from "sonner";

type QuestionType = "discursiva" | "multipla" | "jogo";

type Question = {
  id: number;
  type: QuestionType;
  prompt: string;
  secondaryPrompt: string;
  points: string;
  options: string[];
  image?: string;
  imageWidth: number;
  imageCaption: string;
  gameKind?: string;
  correctOption?: number | null;
};

type ExamData = {
  schoolName: string;
  level: string;
  subject: string;
  teacher: string;
  bimester: string;
  student: string;
  className: string;
  shift: string;
  dateDay: string;
  dateMonth: string;
  title: string;
  notice: string;
  instructions: string;
  extraActivity: string;
  extraActivityLayout: "none" | "single" | "double";
  fontFamily: "Arial" | "Times New Roman";
  fontSize: string;
  antiCheat: boolean;
  variationCount: string;
  shuffleDiscursive: boolean;
  logo?: string;
};

const initialExam: ExamData = {
  schoolName: "",
  level: "",
  subject: "",
  teacher: "",
  bimester: "",
  student: "",
  className: "",
  shift: "",
  dateDay: "",
  dateMonth: "",
  title: "AVALIAÇÃO",
  notice: "",
  instructions: "",
  extraActivity: "",
  extraActivityLayout: "none",
  fontFamily: "Arial",
  fontSize: "12",
  antiCheat: false,
  variationCount: "1",
  shuffleDiscursive: false,
};

const initialQuestions: Question[] = [{ id: 1, type: "discursiva", prompt: "", secondaryPrompt: "", points: "", options: ["", "", "", ""], imageWidth: 100, imageCaption: "" }];

type Variant = { number: number; questions: Question[] };

function seededShuffle<T>(items: T[], seed: number) {
  const output = [...items];
  let value = seed * 9301 + 49297;
  for (let index = output.length - 1; index > 0; index -= 1) {
    value = (value * 233280 + 12345) % 2147483647;
    const target = Math.floor((value / 2147483647) * (index + 1));
    [output[index], output[target]] = [output[target], output[index]];
  }
  return output;
}

function estimatedQuestionHeight(question: Question) {
  const textLength = (question.prompt || "").replace(/<[^>]+>/g, "").length;
  const secondaryLength = (question.secondaryPrompt || "").replace(/<[^>]+>/g, "").length;
  let height = 105 + Math.ceil(textLength / 50) * 17 + Math.ceil(secondaryLength / 50) * 16;
  if (question.image) height += 280;
  if (question.imageCaption) height += 24;
  if (question.type === "multipla") height += question.options.reduce((total, option) => total + 20 + Math.ceil(option.length / 45) * 14, 0);
  if (question.type === "discursiva") height += 78;
  if (question.type === "jogo") height += 40;
  return height;
}

function splitQuestionsIntoPages(items: Question[], mode: "single" | "double") {
  const maximumPerPage = mode === "single" ? 6 : 6;
  const capacity = mode === "single" ? 1160 : 1450;
  const pages: Question[][] = [];
  let current: Question[] = [];
  let currentHeight = 0;
  items.forEach((question) => {
    const height = estimatedQuestionHeight(question);
    const candidate = [...current, question];
    const candidateHeight = mode === "double"
      ? Math.max(
        candidate.slice(0, Math.ceil(candidate.length / 2)).reduce((total, item) => total + estimatedQuestionHeight(item), 0),
        candidate.slice(Math.ceil(candidate.length / 2)).reduce((total, item) => total + estimatedQuestionHeight(item), 0),
      )
      : currentHeight + height;
    if (current.length && (current.length >= maximumPerPage || candidateHeight > capacity)) {
      pages.push(current);
      current = [];
      currentHeight = 0;
    }
    current.push(question);
    currentHeight += height;
  });
  if (current.length || !pages.length) pages.push(current);
  return pages;
}

const STORAGE_KEY = "prova-pronta-draft-v4";

function nextQuestionId(questions: Question[]) {
  return questions.length ? Math.max(...questions.map((question) => question.id)) + 1 : 1;
}

function Field({
  label,
  value,
  onChange,
  placeholder,
  className = "",
  type = "text",
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  className?: string;
  type?: string;
}) {
  return (
    <label className={`field ${className}`}>
      <span>{label}</span>
      <input type={type} value={value} placeholder={placeholder} onChange={(event) => onChange(event.target.value)} />
    </label>
  );
}

function RichTextField({ value, onChange, placeholder }: { value: string; onChange: (value: string) => void; placeholder: string }) {
  const editorRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (editorRef.current && editorRef.current.innerHTML !== value) editorRef.current.innerHTML = value;
  }, [value]);
  const command = (name: string, commandValue?: string) => {
    editorRef.current?.focus();
    document.execCommand(name, false, commandValue);
    onChange(editorRef.current?.innerHTML || "");
  };
  return (
    <div className="rich-field">
      <div className="rich-toolbar" aria-label="Formatação do texto">
        <button type="button" onMouseDown={(event) => event.preventDefault()} onClick={() => command("bold")}><b>B</b></button>
        <button type="button" onMouseDown={(event) => event.preventDefault()} onClick={() => command("italic")}><i>I</i></button>
        <button type="button" onMouseDown={(event) => event.preventDefault()} onClick={() => command("underline")}><u>U</u></button>
        <span />
        <button type="button" onMouseDown={(event) => event.preventDefault()} onClick={() => command("justifyLeft")}>≡</button>
        <button type="button" onMouseDown={(event) => event.preventDefault()} onClick={() => command("justifyCenter")}>≡</button>
        <button type="button" onMouseDown={(event) => event.preventDefault()} onClick={() => command("justifyFull")} title="Justificar">☰</button>
        <button type="button" onMouseDown={(event) => event.preventDefault()} onClick={() => command("insertUnorderedList")}>•</button>
      </div>
      <div ref={editorRef} className="rich-editor" contentEditable suppressContentEditableWarning data-placeholder={placeholder} onInput={(event) => onChange(event.currentTarget.innerHTML)} />
    </div>
  );
}

function SectionHeader({ number, eyebrow, title, description }: { number: string; eyebrow: string; title: string; description: string }) {
  return (
    <div className="section-header">
      <div className="section-number">{number}</div>
      <div>
        <div className="eyebrow">{eyebrow}</div>
        <h2>{title}</h2>
        <p>{description}</p>
      </div>
    </div>
  );
}

function LogoMark({ small = false }: { small?: boolean }) {
  return (
    <div className={`logo-mark ${small ? "logo-mark-small" : ""}`} aria-hidden="true">
      <span className="logo-line logo-line-one" />
      <span className="logo-line logo-line-two" />
      <span className="logo-dot" />
    </div>
  );
}

function QuestionPreview({ question, index }: { question: Question; index: number }) {
  return (
    <div className="question-preview">
      <div className="question-title">
        <strong>{index + 1}.</strong>
        <span dangerouslySetInnerHTML={{ __html: question.prompt || "" }} />
        {question.points && <em>({question.points} pt)</em>}
      </div>
      {question.image && <img className="question-image" src={question.image} alt="Imagem da questão" style={{ width: `${question.imageWidth || 100}%` }} />}
      {question.image && question.imageCaption && <div className="question-caption">{question.imageCaption}</div>}
      {question.secondaryPrompt && <div className="question-secondary" dangerouslySetInnerHTML={{ __html: question.secondaryPrompt }} />}
      {question.type === "jogo" ? null : question.type === "multipla" ? (
        <div className="options-preview">
          {question.options.map((option, optionIndex) => (
            <div className="option-preview" key={`${question.id}-${optionIndex}`}>
              <span className="option-circle">{String.fromCharCode(65 + optionIndex)}</span>
              <span>{option || "Alternativa"}</span>
            </div>
          ))}
        </div>
      ) : (
        <div className="answer-lines">
          <span />
          <span />
          <span />
        </div>
      )}
    </div>
  );
}

function PaperPage({ exam, logoPreview, pageMode, variantNumber, questions, startIndex, extraActivity, extraActivityLayout, pageNumber, pageCount }: { exam: ExamData; logoPreview: string; pageMode: "single" | "double"; variantNumber: number; questions: Question[]; startIndex: number; extraActivity?: string; extraActivityLayout?: "none" | "single" | "double"; pageNumber: number; pageCount: number }) {
  const isExtra = Boolean(extraActivity);
  return <div className={`paper ${pageNumber > 1 ? "paper-continuation" : ""}`}>
    {!isExtra && pageNumber === 1 && <div className="paper-header-table">
      <div className="paper-logo-cell">{logoPreview ? <img src={logoPreview} alt="Logo" /> : <div className="logo-placeholder">LOGO<br />DA ESCOLA</div>}</div>
      <div className="paper-school-cell">{exam.schoolName}{exam.schoolName && exam.level ? ` – ${exam.level.toUpperCase()}` : exam.level.toUpperCase()}</div>
      <div className="paper-info-row"><span><b>DISCIPLINA:</b> {exam.subject || "________________"}</span><span><b>PROFESSOR(A):</b> {exam.teacher || "________________"}</span><span><b>BIMESTRE:</b> {exam.bimester ? exam.bimester.toUpperCase() : "____"}</span></div>
      <div className="paper-student-row"><span><b>ALUNO(A):</b> {exam.student || "____________________________________________________________"}</span></div>
      <div className="paper-info-row"><span><b>ANO:</b> {exam.className || "____"}</span><span><b>TURMA:</b> __________</span><span><b>TURNO:</b> {exam.shift || "____"}</span><span><b>DATA:</b> {exam.dateDay || "____"}/{exam.dateMonth || "____"}/2026</span></div>
    </div>}
    {!isExtra && pageNumber === 1 && <div className="paper-title-block" style={{ fontFamily: exam.fontFamily, fontSize: `${exam.fontSize}px` }}><h1>{exam.title || "AVALIAÇÃO"}</h1>{exam.notice && <div className="paper-notice" dangerouslySetInnerHTML={{ __html: exam.notice }} />}{exam.instructions && <p dangerouslySetInnerHTML={{ __html: exam.instructions }} />}</div>}
    {isExtra && <div className="paper-title-block extra-activity-title"><h1>ATIVIDADE EXTRA</h1></div>}
    {isExtra ? <div className={`extra-activity-content ${extraActivityLayout === "single" ? "single-column" : "double-column"}`} dangerouslySetInnerHTML={{ __html: extraActivity || "" }} /> : <div className={`paper-columns ${pageMode === "single" ? "single-column" : ""}`} style={{ fontFamily: exam.fontFamily, fontSize: `${exam.fontSize}px` }}>{pageMode === "single" ? questions.map((question, index) => <QuestionPreview question={question} index={startIndex + index} key={question.id} />) : <><div className="question-column">{questions.slice(0, Math.ceil(questions.length / 2)).map((question, index) => <QuestionPreview question={question} index={startIndex + index} key={question.id} />)}</div><div className="question-column">{questions.slice(Math.ceil(questions.length / 2)).map((question, index) => <QuestionPreview question={question} index={startIndex + Math.ceil(questions.length / 2) + index} key={question.id} />)}</div></>}</div>}
    <div className="paper-footer"><span>TIPO {variantNumber}{isExtra ? " · ATIVIDADE EXTRA" : ""}</span><span>PÁGINA {pageNumber} DE {pageCount}</span></div>
  </div>;
}

async function examToParagraphs(question: Question, index: number) {
  const plainPrompt = question.prompt.replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ") || "Escreva o enunciado da questão...";
  const prompt = new Paragraph({
    spacing: { after: 80 },
    children: [
      new TextRun({ text: `${index + 1}. `, bold: true }),
      new TextRun({ text: plainPrompt }),
      ...(question.points ? [new TextRun({ text: ` (${question.points} pt)`, italics: true, color: "68717D" })] : []),
    ],
  });
  const secondaryPrompt = question.secondaryPrompt ? new Paragraph({
    spacing: { after: 80 },
    children: [new TextRun({ text: question.secondaryPrompt.replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ") })],
  }) : null;
  const imageParagraph = question.image ? new Paragraph({
    children: [new ImageRun({ data: await fetch(question.image).then((response) => response.arrayBuffer()), transformation: await imageTransformation(question.image, question.imageWidth || 100), type: question.image.startsWith("data:image/jpeg") ? "jpg" : "png" })],
  }) : null;
  const imageCaption = question.image && question.imageCaption ? new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 80 },
    children: [new TextRun({ text: question.imageCaption, italics: true, color: "68717D", size: 16 })],
  }) : null;
  if (question.type === "multipla") {
    return [prompt, ...(imageParagraph ? [imageParagraph] : []), ...(imageCaption ? [imageCaption] : []), ...(secondaryPrompt ? [secondaryPrompt] : []), ...question.options.map((option, optionIndex) => new Paragraph({
      indent: { left: 340 },
      spacing: { after: 40 },
      children: [new TextRun({ text: `${String.fromCharCode(65 + optionIndex)}) `, bold: true }), new TextRun(option || "Alternativa")],
    }))];
  }
  if (question.type === "jogo") return [prompt, ...(imageParagraph ? [imageParagraph] : []), ...(imageCaption ? [imageCaption] : []), ...(secondaryPrompt ? [secondaryPrompt] : [])];
  return [prompt, ...(imageParagraph ? [imageParagraph] : []), ...(imageCaption ? [imageCaption] : []), ...(secondaryPrompt ? [secondaryPrompt] : []), ...[1, 2, 3].map(() => new Paragraph({
    spacing: { after: 220 },
    border: { bottom: { color: "AAB4BF", style: BorderStyle.SINGLE, size: 4 } },
    children: [new TextRun(" ")],
  }))];
}

async function imageTransformation(source: string, percentage = 100) {
  const image = new Image();
  image.src = source;
  await image.decode();
  const width = Math.max(40, Math.min(520, Math.round(280 * (percentage / 100))));
  const height = Math.max(30, Math.round(width * (image.naturalHeight / image.naturalWidth)));
  return { width, height };
}

export default function Home() {
  const [exam, setExam] = useState<ExamData>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      return saved ? { ...initialExam, ...JSON.parse(saved).exam } : initialExam;
    } catch {
      return initialExam;
    }
  });
  const [questions, setQuestions] = useState<Question[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      return saved ? JSON.parse(saved).questions || initialQuestions : initialQuestions;
    } catch {
      return initialQuestions;
    }
  });
  const [activeSection, setActiveSection] = useState("identidade");
  const [saved, setSaved] = useState(true);
  const [logoPreview, setLogoPreview] = useState(exam.logo || "");
  const [pageMode, setPageMode] = useState<"single" | "double">("double");
  const [activeVariant, setActiveVariant] = useState(1);
  const fileRef = useRef<HTMLInputElement>(null);

  const updateExam = (key: keyof ExamData, value: string | boolean) => {
    setExam((current) => ({ ...current, [key]: value }));
    setSaved(false);
  };

  const variants = useMemo<Variant[]>(() => {
    const count = exam.antiCheat ? Math.min(4, Math.max(1, Number(exam.variationCount) || 1)) : 1;
    return Array.from({ length: count }, (_, index) => {
      const multiple = questions.filter((question) => question.type === "multipla");
      const other = questions.filter((question) => question.type !== "multipla");
      const shuffledMultiple = exam.antiCheat ? seededShuffle(multiple, index + 11).map((question) => {
        const optionOrder = seededShuffle(question.options.map((_, optionIndex) => optionIndex), index * 97 + question.id * 13);
        return { ...question, options: optionOrder.map((optionIndex) => question.options[optionIndex]), correctOption: question.correctOption === undefined || question.correctOption === null ? question.correctOption : optionOrder.indexOf(question.correctOption) };
      }) : multiple;
      const shuffledOther = exam.antiCheat && exam.shuffleDiscursive ? seededShuffle(other, index + 71) : other;
      return { number: index + 1, questions: exam.antiCheat ? [...shuffledMultiple, ...shuffledOther] : questions };
    });
  }, [exam.antiCheat, exam.variationCount, exam.shuffleDiscursive, questions]);

  const previewQuestions = variants[activeVariant - 1]?.questions || questions;
  const previewQuestionPages = useMemo(() => splitQuestionsIntoPages(previewQuestions, pageMode), [previewQuestions, pageMode]);
  useEffect(() => {
    if (activeVariant > variants.length) setActiveVariant(1);
  }, [activeVariant, variants.length]);

  const updateQuestion = (id: number, patch: Partial<Question>) => {
    setQuestions((current) => current.map((question) => (question.id === id ? { ...question, ...patch } : question)));
    setSaved(false);
  };

  const updateOption = (id: number, optionIndex: number, value: string) => {
    setQuestions((current) => current.map((question) => {
      if (question.id !== id) return question;
      return { ...question, options: question.options.map((option, index) => (index === optionIndex ? value : option)) };
    }));
    setSaved(false);
  };

  const addQuestion = () => {
    if (questions.length >= 10) {
      toast.error("Você já atingiu o limite de 10 questões.");
      return;
    }
    const id = nextQuestionId(questions);
    setQuestions((current) => [...current, { id, type: "discursiva", prompt: "", secondaryPrompt: "", points: "", options: ["", "", "", ""], imageWidth: 100, imageCaption: "" }]);
    setSaved(false);
    setTimeout(() => document.getElementById(`question-${id}`)?.scrollIntoView({ behavior: "smooth", block: "center" }), 50);
  };

  const deleteQuestion = (id: number) => {
    setQuestions((current) => current.filter((question) => question.id !== id));
    setSaved(false);
  };

  const moveQuestion = (index: number, direction: -1 | 1) => {
    const destination = index + direction;
    if (destination < 0 || destination >= questions.length) return;
    const next = [...questions];
    [next[index], next[destination]] = [next[destination], next[index]];
    setQuestions(next);
    setSaved(false);
  };

  const saveDraft = () => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ exam: { ...exam, logo: logoPreview }, questions }));
    setSaved(true);
    toast.success("Rascunho salvo neste navegador.");
  };

  const resetDraft = () => {
    setExam(initialExam);
    setQuestions(initialQuestions);
    setLogoPreview("");
    setPageMode("double");
    localStorage.removeItem(STORAGE_KEY);
    setSaved(true);
    toast.success("Modelo inicial restaurado.");
  };

  useEffect(() => {
    const timer = window.setTimeout(() => {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ exam: { ...exam, logo: logoPreview }, questions }));
      setSaved(true);
    }, 700);
    return () => window.clearTimeout(timer);
  }, [exam, questions, logoPreview]);

  const handleLogo = (file?: File) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      setLogoPreview(result);
      setSaved(false);
    };
    reader.readAsDataURL(file);
  };

  const handleQuestionImage = (id: number, file?: File) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => updateQuestion(id, { image: String(reader.result || "") });
    reader.readAsDataURL(file);
  };

  const exportPdf = () => {
    document.body.classList.add("printing-exam");
    window.setTimeout(() => {
      window.print();
      window.setTimeout(() => document.body.classList.remove("printing-exam"), 400);
    }, 50);
  };

  const printAnswerKey = () => {
    document.body.classList.add("printing-answer-key");
    window.setTimeout(() => {
      window.print();
      window.setTimeout(() => document.body.classList.remove("printing-answer-key"), 400);
    }, 50);
  };

  const downloadAnswerKey = async () => {
    const paragraphs = variants.map((variant) => [
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: `TIPO ${variant.number}`, bold: true })] }),
      ...variant.questions.filter((question) => question.type === "multipla").map((question) => {
        const variantIndex = variant.questions.findIndex((item) => item.id === question.id);
        const answer = question.correctOption === undefined || question.correctOption === null ? "não informado" : String.fromCharCode(65 + question.correctOption);
        return new Paragraph({ children: [new TextRun({ text: `${variantIndex + 1}. ${answer}`, bold: true }), new TextRun({ text: question.points ? ` — ${question.points} ponto(s)` : "" })] });
      }),
      new Paragraph({ children: [new TextRun({ text: " " })] }),
    ]).flat();
    const doc = new Document({ sections: [{ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "GABARITO", bold: true, size: 28 })] }), new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: exam.title || "AVALIAÇÃO", size: 20 })] }), ...paragraphs] }] });
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "gabarito.docx";
    link.click();
    URL.revokeObjectURL(url);
    toast.success("Gabarito DOCX baixado.");
  };

  const exportDocx = async () => {
    const logoType = logoPreview.startsWith("data:image/jpeg") || logoPreview.startsWith("data:image/jpg") ? "jpg" : "png";
    const headerRows = [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 2200, type: WidthType.DXA },
            children: [logoPreview ? new Paragraph({ children: [new ImageRun({ data: await fetch(logoPreview).then((response) => response.arrayBuffer()), transformation: { width: 54, height: 54 }, type: logoType })] }) : new Paragraph({ children: [new TextRun({ text: "PROVA", bold: true, color: "193A58", size: 20 })] })],
          }),
          new TableCell({
            width: { size: 7000, type: WidthType.DXA },
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: exam.schoolName || "Nome da escola", bold: true, size: 25, color: "193A58" })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${exam.level || "Nível de ensino"} · ${exam.subject || "Disciplina"}`, size: 18, color: "68717D" })] }),
            ],
          }),
        ],
      }),
    ];
    const details = new Table({
      width: { size: 9200, type: WidthType.DXA },
      rows: [
        ...headerRows,
        new TableRow({ children: [new TableCell({ columnSpan: 2, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 100 }, children: [new TextRun({ text: exam.title || "Título da avaliação", bold: true, size: 28, color: "193A58" })] })] })] }),
        new TableRow({ children: [new TableCell({ columnSpan: 2, children: [new Paragraph({ children: [new TextRun({ text: `Aluno(a): ${exam.student || "________________________________________"}`, size: 18 }), new TextRun({ text: `    Turma: ${exam.className || "________"}`, size: 18 }), new TextRun({ text: `    Turno: ${exam.shift || "________"}`, size: 18 })] }), new Paragraph({ children: [new TextRun({ text: `Professor(a): ${exam.teacher || "____________________________"}`, size: 18 }), new TextRun({ text: `    ${exam.bimester || "Bimestre"}`, size: 18 }), new TextRun({ text: `    Data: ____ / ____`, size: 18 })] })] })] }),
      ],
    });
    const paragraphs = (await Promise.all(previewQuestions.map(examToParagraphs))).flat();
    const doc = new Document({
      sections: [{
        properties: { page: { margin: { top: 520, right: 520, bottom: 520, left: 520 } } },
        children: [
          details,
          ...(exam.notice ? [new Paragraph({ spacing: { before: 120, after: 90 }, children: [new TextRun({ text: exam.notice.replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " "), size: Number(exam.fontSize) * 2 })] })] : []),
          ...(exam.instructions ? [new Paragraph({ spacing: { before: 90, after: 120 }, children: [new TextRun({ text: exam.instructions.replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " "), italics: true, color: "68717D", size: Number(exam.fontSize) * 2 })] })] : []),
          new Table({
            width: { size: 9200, type: WidthType.DXA },
            rows: [new TableRow({ children: pageMode === "single" ? [new TableCell({ width: { size: 9200, type: WidthType.DXA }, children: paragraphs })] : [new TableCell({ width: { size: 4600, type: WidthType.DXA }, children: paragraphs.filter((_, index) => index % 2 === 0) }), new TableCell({ width: { size: 4600, type: WidthType.DXA }, children: paragraphs.filter((_, index) => index % 2 === 1) })] })],
          }),
          new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 200 }, children: [new TextRun({ text: `TIPO ${activeVariant} · ${previewQuestions.length} questão${previewQuestions.length === 1 ? "" : "ões"}`, italics: true, color: "68717D", size: 16 })] }),
        ],
      }],
    });
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${(exam.title || "prova").toLowerCase().replace(/[^a-z0-9]+/gi, "-")}.docx`;
    link.click();
    URL.revokeObjectURL(url);
    toast.success("Arquivo DOCX baixado.");
  };

  const sectionTabs = [
    { id: "identidade", label: "Identidade", icon: FileText },
    { id: "questoes", label: "Questões", icon: AlignJustify },
    { id: "anticola", label: "Anti-cola", icon: ShieldCheck },
    { id: "visual", label: "Visual", icon: LayoutGrid },
  ];

  const questionCountLabel = useMemo(() => `${questions.length}/10 questões`, [questions.length]);

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand-lockup">
          <LogoMark />
          <div>
            <div className="brand-name">prova<span>pronta</span></div>
            <div className="brand-caption">editor de avaliações</div>
          </div>
        </div>
        <div className="topbar-meta">
          <div className="save-state"><span className={`status-dot ${saved ? "is-saved" : ""}`} />{saved ? "Salvo automaticamente" : "Alterações pendentes"}</div>
          <button className="icon-button" title="Ajuda"><CircleHelp size={18} /></button>
          <div className="avatar">MA</div>
        </div>
      </header>

      <main className="workspace">
        <aside className="editor-panel">
          <div className="editor-heading">
            <div>
              <div className="eyebrow">NOVO DOCUMENTO</div>
              <h1>Monte sua prova</h1>
              <p>Preencha os campos e acompanhe o resultado na prévia ao lado.</p>
            </div>
            <button className="more-button" title="Mais opções"><Settings2 size={18} /></button>
          </div>

          <nav className="section-tabs" aria-label="Etapas da prova">
            {sectionTabs.map(({ id, label, icon: Icon }) => (
              <button key={id} className={activeSection === id ? "active" : ""} onClick={() => { setActiveSection(id); document.getElementById(id)?.scrollIntoView({ behavior: "smooth", block: "start" }); }}>
                <Icon size={16} /> <span>{label}</span>
                {id === "questoes" && <small>{questions.length}</small>}
              </button>
            ))}
          </nav>

          <div className="editor-scroll">
            <section className="editor-section" id="identidade">
              <SectionHeader number="01" eyebrow="IDENTIDADE" title="A cara da sua escola" description="Defina as informações que aparecem no topo da avaliação." />
              <div className="logo-upload-row">
                <button className="logo-uploader" onClick={() => fileRef.current?.click()} aria-label="Adicionar logo da escola">
                  {logoPreview ? <img src={logoPreview} alt="Logo da escola" /> : <ImagePlus size={22} />}
                  {!logoPreview && <span>Logo</span>}
                </button>
                <div className="logo-copy"><strong>Logo da escola</strong><span>PNG ou JPG · opcional</span></div>
                {logoPreview && <button className="remove-logo" onClick={() => setLogoPreview("")}><X size={14} /></button>}
                <input ref={fileRef} className="hidden-input" type="file" accept="image/png,image/jpeg" onChange={(event) => handleLogo(event.target.files?.[0])} />
              </div>
              <div className="field-grid two-columns">
                <Field label="Nome da escola" value={exam.schoolName} onChange={(value) => updateExam("schoolName", value)} placeholder="Ex.: Colégio Horizonte" />
                <label className="field"><span>Nível de ensino</span><select value={exam.level} onChange={(event) => updateExam("level", event.target.value)}><option value="">Selecione...</option><option>Educação Infantil</option><option>Ensino Fundamental I</option><option>Ensino Fundamental II</option><option>Ensino Médio</option></select></label>
                <Field label="Disciplina" value={exam.subject} onChange={(value) => updateExam("subject", value)} placeholder="Ex.: Ciências" />
                <Field label="Nome do professor(a)" value={exam.teacher} onChange={(value) => updateExam("teacher", value)} placeholder="Ex.: Prof.ª Marina" />
              </div>
              <div className="field-grid three-columns">
                <label className="field"><span>Bimestre</span><select value={exam.bimester} onChange={(event) => updateExam("bimester", event.target.value)}><option value="">Selecione...</option><option>1º bimestre</option><option>2º bimestre</option><option>3º bimestre</option><option>4º bimestre</option></select></label>
                <Field label="Turma" value={exam.className} onChange={(value) => updateExam("className", value)} placeholder="8º A" />
                <label className="field"><span>Turno</span><select value={exam.shift} onChange={(event) => updateExam("shift", event.target.value)}><option value="">Selecione...</option><option>Manhã</option><option>Tarde</option><option>Noite</option><option>Integral</option></select></label>
              </div>
              <div className="field-grid three-columns date-fields">
                <Field label="Dia" value={exam.dateDay} onChange={(value) => updateExam("dateDay", value)} placeholder="____" />
                <Field label="Mês" value={exam.dateMonth} onChange={(value) => updateExam("dateMonth", value)} placeholder="____" />
                <div className="date-hint"><Info size={14} /> deixe em branco para preencher à caneta</div>
              </div>
              <Field label="Nome do aluno(a)" value={exam.student} onChange={(value) => updateExam("student", value)} placeholder="Será deixado um espaço para preenchimento" className="full-field" />
            </section>

            <section className="editor-section" id="questoes">
              <SectionHeader number="02" eyebrow="CONTEÚDO" title="As questões" description="Escreva até 10 questões. Você pode combinar perguntas abertas e múltipla escolha." />
              <div className="question-limit"><div><strong>{questionCountLabel}</strong><span> · a prova fica mais objetiva assim.</span></div><div className="limit-bar"><span style={{ width: `${questions.length * 10}%` }} /></div></div>
              <div className="question-list">
                {questions.map((question, index) => (
                  <article className="question-card" id={`question-${question.id}`} key={question.id}>
                    <div className="question-card-top">
                      <div className="drag-handle"><AlignJustify size={17} /><strong>Questão {String(index + 1).padStart(2, "0")}</strong></div>
                      <div className="question-actions">
                        <button title="Mover para cima" onClick={() => moveQuestion(index, -1)} disabled={index === 0}><ArrowUp size={15} /></button>
                        <button title="Mover para baixo" onClick={() => moveQuestion(index, 1)} disabled={index === questions.length - 1}><ArrowDown size={15} /></button>
                        <button className="delete-action" title="Excluir questão" onClick={() => deleteQuestion(question.id)}><Trash2 size={15} /></button>
                      </div>
                    </div>
                    <div className="question-card-body">
                      <div className="question-controls">
                        <label className="field"><span>Tipo de questão</span><select value={question.type} onChange={(event) => updateQuestion(question.id, { type: event.target.value as QuestionType })}><option value="discursiva">Discursiva</option><option value="multipla">Múltipla escolha</option><option value="jogo">Atividade / jogo</option></select></label>
                        <Field label="Valor (opcional)" value={question.points} onChange={(value) => updateQuestion(question.id, { points: value })} placeholder="—" />
                      </div>
                      {question.type === "jogo" && <label className="field game-kind-field"><span>Tipo de atividade</span><select value={question.gameKind || "Caça-palavras"} onChange={(event) => updateQuestion(question.id, { gameKind: event.target.value })}><option>Caça-palavras</option><option>Cruzadinha</option><option>Palavras embaralhadas</option><option>Jogo da memória</option><option>Outra atividade</option></select></label>}
                      <div className="field"><span>Enunciado</span><RichTextField value={question.prompt} onChange={(value) => updateQuestion(question.id, { prompt: value })} placeholder="Digite o enunciado da questão..." /></div>
                      <div className="field"><span>Continuação do enunciado principal ou novo enunciado (opcional)</span><RichTextField value={question.secondaryPrompt || ""} onChange={(value) => updateQuestion(question.id, { secondaryPrompt: value })} placeholder="Digite a continuação do enunciado principal ou um novo enunciado..." /></div>
                      <div className="question-media-row"><label className="image-question-button"><ImagePlus size={15} /> {question.image ? "Trocar imagem" : "Inserir imagem"}<input type="file" accept="image/png,image/jpeg,image/webp" onChange={(event) => handleQuestionImage(question.id, event.target.files?.[0])} /></label>{question.image && <button type="button" className="remove-image-button" onClick={() => updateQuestion(question.id, { image: "" })}><X size={13} /> Remover</button>}</div>
                      {question.image && <label className="image-size-control"><span>Tamanho da imagem</span><input type="range" min="30" max="180" step="5" value={question.imageWidth || 100} onChange={(event) => updateQuestion(question.id, { imageWidth: Number(event.target.value) })} /><strong>{question.imageWidth || 100}%</strong></label>}
                      {question.image && <Field label="Legenda da imagem (opcional)" value={question.imageCaption || ""} onChange={(value) => updateQuestion(question.id, { imageCaption: value })} placeholder="Fonte, autor ou informação complementar" />}
                      {question.image && <img className="editor-question-image" src={question.image} alt="Prévia da imagem da questão" />}
                      {question.type === "multipla" && <div className="options-editor"><div className="field-label">Alternativas</div>{question.options.map((option, optionIndex) => <div className="option-row" key={optionIndex}><span>{String.fromCharCode(65 + optionIndex)}</span><input value={option} placeholder={`Alternativa ${String.fromCharCode(65 + optionIndex)}`} onChange={(event) => updateOption(question.id, optionIndex, event.target.value)} /><button type="button" className={`correct-option ${question.correctOption === optionIndex ? "selected" : ""}`} title="Marcar alternativa correta" onClick={() => updateQuestion(question.id, { correctOption: question.correctOption === optionIndex ? null : optionIndex })}>{question.correctOption === optionIndex ? <Check size={12} /> : "✓"}</button></div>)}</div>}
                    </div>
                  </article>
                ))}
              </div>
              <button className="add-question" onClick={addQuestion} disabled={questions.length >= 10}><Plus size={17} /> Adicionar questão <span>{questions.length >= 10 ? "limite atingido" : `${10 - questions.length} restantes`}</span></button>
            </section>

            <section className="editor-section" id="anticola">
              <SectionHeader number="03" eyebrow="SEGURANÇA" title="Anti-cola" description="Crie versões diferentes da prova sem alterar o conteúdo das questões." />
              <div className="anti-cheat-card">
                <div className="anti-cheat-card-title"><div className="anti-cheat-icon"><ShieldCheck size={20} /></div><div><strong>Ativar anti-cola</strong><span>Embaralha questões de múltipla escolha e suas alternativas.</span></div><label className="switch"><input type="checkbox" checked={exam.antiCheat} onChange={(event) => updateExam("antiCheat", event.target.checked)} /><span /></label></div>
                {exam.antiCheat && <div className="anti-cheat-settings"><label className="field"><span>Quantidade de tipos de prova</span><select value={exam.variationCount} onChange={(event) => { updateExam("variationCount", event.target.value); setActiveVariant(1); }}><option value="1">1 tipo</option><option value="2">2 tipos</option><option value="3">3 tipos</option><option value="4">4 tipos</option></select></label><label className="check-row"><input type="checkbox" checked={exam.shuffleDiscursive} onChange={(event) => updateExam("shuffleDiscursive", event.target.checked)} /><span>Também reorganizar a ordem das questões discursivas</span></label></div>}
              </div>
              <div className="variant-explanation"><KeyRound size={16} /><div><strong>Como funciona</strong><p>As questões objetivas e as alternativas são reorganizadas por tipo. Cada versão recebe um número e possui seu próprio gabarito.</p></div></div>
            </section>

            <section className="editor-section" id="visual">
              <SectionHeader number="04" eyebrow="ACABAMENTO" title="Um toque final" description="Pequenos detalhes que deixam sua avaliação pronta para imprimir." />
              <Field label="Título da avaliação" value={exam.title} onChange={(value) => updateExam("title", value)} placeholder="Ex.: AVALIAÇÃO DO 3º BIMESTRE" />
              <div className="field rich-field-wrap"><span>Avisos, textos ou conteúdo complementar</span><RichTextField value={exam.notice} onChange={(value) => updateExam("notice", value)} placeholder="Digite aqui um texto, aviso, texto-base ou instruções..." /></div>
              <div className="field rich-field-wrap"><span>Instruções para a turma</span><RichTextField value={exam.instructions} onChange={(value) => updateExam("instructions", value)} placeholder="Ex.: Leia cada questão com atenção..." /></div>
              <div className="field-grid two-columns"><label className="field"><span>Fonte da prova</span><select value={exam.fontFamily} onChange={(event) => updateExam("fontFamily", event.target.value as ExamData["fontFamily"])}><option>Arial</option><option>Times New Roman</option></select></label><label className="field"><span>Tamanho da fonte</span><select value={exam.fontSize} onChange={(event) => updateExam("fontSize", event.target.value)}>{["8", "9", "10", "11", "12", "13", "14", "15", "16"].map((size) => <option key={size} value={size}>{size} pt</option>)}</select></label></div>
              <div className="field rich-field-wrap"><span>Atividade extra (opcional)</span><RichTextField value={exam.extraActivity} onChange={(value) => updateExam("extraActivity", value)} placeholder="Insira o texto da atividade extra ou use a imagem da atividade em uma questão..." /></div>
              <label className="field"><span>Distribuição da atividade extra</span><select value={exam.extraActivityLayout} onChange={(event) => updateExam("extraActivityLayout", event.target.value as ExamData["extraActivityLayout"])}><option value="none">Não inserir atividade extra</option><option value="single">Uma coluna</option><option value="double">Duas colunas · frente e verso</option></select></label>
              <div className="paper-mode-control"><span>Distribuição da folha</span><div><button className={pageMode === "single" ? "selected" : ""} onClick={() => setPageMode("single")}>1 lado · uma coluna</button><button className={pageMode === "double" ? "selected" : ""} onClick={() => setPageMode("double")}>2 lados · duas colunas</button></div></div>
              <div className="tip-card"><Sparkles size={18} /><div><strong>Atividades diferentes também cabem aqui</strong><p>Use o tipo “Atividade / jogo” para reservar um espaço para caça-palavras, cruzadinha ou outro material.</p></div></div>
              <button className="reset-button" onClick={resetDraft}><RotateCcw size={15} /> Restaurar modelo inicial</button>
            </section>
          </div>
          <div className="editor-footer"><button className="save-button" onClick={saveDraft}><Save size={16} /> Salvar rascunho</button><span>As alterações ficam neste dispositivo.</span></div>
        </aside>

        <section className="preview-panel">
          <div className="preview-toolbar">
            <div><div className="eyebrow">PRÉVIA DA FOLHA</div><h2>Veja antes de baixar</h2><div className="variant-picker"><Shuffle size={13} /> {variants.map((variant) => <button key={variant.number} className={activeVariant === variant.number ? "selected" : ""} onClick={() => setActiveVariant(variant.number)}>Tipo {variant.number}</button>)}</div></div>
            <div className="preview-actions"><button className="secondary-button" onClick={exportPdf}><Printer size={16} /> Exportar PDF</button><button className="primary-button" onClick={exportDocx}><FileDown size={16} /> Baixar DOCX</button><button className="key-button" onClick={downloadAnswerKey}><KeyRound size={16} /> Gabarito DOCX</button><button className="key-button" onClick={printAnswerKey}><Printer size={16} /> Gabarito PDF</button></div>
          </div>
          <div className="preview-stage">
            <div id="paper-preview" className="paper-set">
              {previewQuestionPages.map((pageQuestions, pageIndex) => <PaperPage key={`active-${activeVariant}-${pageIndex}`} exam={exam} logoPreview={logoPreview} pageMode={pageMode} variantNumber={activeVariant} questions={pageQuestions} startIndex={previewQuestions.indexOf(pageQuestions[0])} pageNumber={pageIndex + 1} pageCount={previewQuestionPages.length} />)}
              {exam.extraActivity && <PaperPage exam={exam} logoPreview={logoPreview} pageMode={pageMode} variantNumber={activeVariant} questions={[]} startIndex={0} extraActivity={exam.extraActivity} extraActivityLayout={exam.extraActivityLayout} pageNumber={previewQuestionPages.length + 1} pageCount={previewQuestionPages.length + 1} />}
              {variants.filter((variant) => variant.number !== activeVariant).map((variant) => {
                const pages = splitQuestionsIntoPages(variant.questions, pageMode);
                return <div className="print-variant" key={`print-variant-${variant.number}`}>{pages.map((pageQuestions, pageIndex) => <PaperPage key={`${variant.number}-${pageIndex}`} exam={exam} logoPreview={logoPreview} pageMode={pageMode} variantNumber={variant.number} questions={pageQuestions} startIndex={variant.questions.indexOf(pageQuestions[0])} pageNumber={pageIndex + 1} pageCount={pages.length} />)}{exam.extraActivity && <PaperPage exam={exam} logoPreview={logoPreview} pageMode={pageMode} variantNumber={variant.number} questions={[]} startIndex={0} extraActivity={exam.extraActivity} extraActivityLayout={exam.extraActivityLayout} pageNumber={pages.length + 1} pageCount={pages.length + 1} />}</div>;
              })}
            </div>
            <div className="answer-key-print"><h1>GABARITO</h1><h2>{exam.title || "AVALIAÇÃO"}</h2>{variants.map((variant) => <section key={variant.number}><h3>TIPO {variant.number}</h3>{variant.questions.filter((question) => question.type === "multipla").map((question) => <p key={question.id}>{variant.questions.findIndex((item) => item.id === question.id) + 1}. <strong>{question.correctOption === undefined || question.correctOption === null ? "não informado" : String.fromCharCode(65 + question.correctOption)}</strong></p>)}</section>)}</div>
          </div>
          <div className="preview-note"><Check size={15} /><span>Formato A4 · {pageMode === "single" ? "uma coluna" : "duas colunas"} · {previewQuestionPages.length + (exam.extraActivity ? 1 : 0)} página{previewQuestionPages.length + (exam.extraActivity ? 1 : 0) === 1 ? "" : "s"} na prévia{exam.antiCheat && variants.length > 1 ? ` · PDF com ${variants.length} tipos` : ""} · fonte {exam.fontFamily}, {exam.fontSize} pt</span><ChevronDown size={15} /></div>
        </section>
      </main>
    </div>
  );
}
