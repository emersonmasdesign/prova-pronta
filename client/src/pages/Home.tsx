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

type QuestionType = "discursiva" | "multipla";

type Question = {
  id: number;
  type: QuestionType;
  prompt: string;
  points: string;
  options: string[];
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
  instructions: string;
  logo?: string;
};

const initialExam: ExamData = {
  schoolName: "Colégio Horizonte",
  level: "Ensino Fundamental II",
  subject: "Ciências",
  teacher: "Prof.ª Marina Alves",
  bimester: "3º bimestre",
  student: "",
  className: "8º ano A",
  shift: "Manhã",
  dateDay: "",
  dateMonth: "",
  title: "AVALIAÇÃO DO 3º BIMESTRE",
  instructions: "Leia cada questão com atenção e responda com clareza.",
};

const initialQuestions: Question[] = [
  {
    id: 1,
    type: "discursiva",
    prompt: "Explique, com suas palavras, por que a preservação da água é importante para a vida no planeta.",
    points: "1,0",
    options: ["", "", "", ""],
  },
  {
    id: 2,
    type: "multipla",
    prompt: "Qual alternativa apresenta uma fonte de energia renovável?",
    points: "1,0",
    options: ["Carvão mineral", "Petróleo", "Energia solar", "Gás natural"],
  },
  {
    id: 3,
    type: "discursiva",
    prompt: "Observe o conteúdo estudado em aula e relacione-o a uma situação do cotidiano.",
    points: "1,0",
    options: ["", "", "", ""],
  },
];

const STORAGE_KEY = "prova-pronta-draft";

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
        <span>{question.prompt || "Escreva o enunciado da questão..."}</span>
        <em>({question.points || "—"} pt)</em>
      </div>
      {question.type === "multipla" ? (
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

function examToParagraphs(question: Question, index: number) {
  const prompt = new Paragraph({
    spacing: { after: 80 },
    children: [
      new TextRun({ text: `${index + 1}. `, bold: true }),
      new TextRun({ text: question.prompt || "Escreva o enunciado da questão..." }),
      new TextRun({ text: ` (${question.points || "—"} pt)`, italics: true, color: "68717D" }),
    ],
  });
  if (question.type === "multipla") {
    return [prompt, ...question.options.map((option, optionIndex) => new Paragraph({
      indent: { left: 340 },
      spacing: { after: 40 },
      children: [new TextRun({ text: `${String.fromCharCode(65 + optionIndex)}) `, bold: true }), new TextRun(option || "Alternativa")],
    }))];
  }
  return [prompt, ...[1, 2, 3].map(() => new Paragraph({
    spacing: { after: 220 },
    border: { bottom: { color: "AAB4BF", style: BorderStyle.SINGLE, size: 4 } },
    children: [new TextRun(" ")],
  }))];
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
  const fileRef = useRef<HTMLInputElement>(null);

  const updateExam = (key: keyof ExamData, value: string) => {
    setExam((current) => ({ ...current, [key]: value }));
    setSaved(false);
  };

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
    setQuestions((current) => [...current, { id, type: "discursiva", prompt: "", points: "1,0", options: ["", "", "", ""] }]);
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

  const exportPdf = () => {
    document.body.classList.add("printing-exam");
    window.setTimeout(() => {
      window.print();
      window.setTimeout(() => document.body.classList.remove("printing-exam"), 400);
    }, 50);
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
    const paragraphs = questions.flatMap(examToParagraphs);
    const doc = new Document({
      sections: [{
        properties: { page: { margin: { top: 520, right: 520, bottom: 520, left: 520 } } },
        children: [
          details,
          new Paragraph({ spacing: { before: 120, after: 120 }, children: [new TextRun({ text: exam.instructions || "", italics: true, color: "68717D", size: 17 })] }),
          new Table({
            width: { size: 9200, type: WidthType.DXA },
            rows: [new TableRow({ children: [new TableCell({ width: { size: 4600, type: WidthType.DXA }, children: paragraphs.filter((_, index) => index % 2 === 0) }), new TableCell({ width: { size: 4600, type: WidthType.DXA }, children: paragraphs.filter((_, index) => index % 2 === 1) })] })],
          }),
          new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 200 }, children: [new TextRun({ text: `${questions.length} questão${questions.length === 1 ? "" : "ões"} · boa prova!`, italics: true, color: "68717D", size: 16 })] }),
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
                <label className="field"><span>Nível de ensino</span><select value={exam.level} onChange={(event) => updateExam("level", event.target.value)}><option>Educação Infantil</option><option>Ensino Fundamental I</option><option>Ensino Fundamental II</option><option>Ensino Médio</option></select></label>
                <Field label="Disciplina" value={exam.subject} onChange={(value) => updateExam("subject", value)} placeholder="Ex.: Ciências" />
                <Field label="Nome do professor(a)" value={exam.teacher} onChange={(value) => updateExam("teacher", value)} placeholder="Ex.: Prof.ª Marina" />
              </div>
              <div className="field-grid three-columns">
                <label className="field"><span>Bimestre</span><select value={exam.bimester} onChange={(event) => updateExam("bimester", event.target.value)}><option>1º bimestre</option><option>2º bimestre</option><option>3º bimestre</option><option>4º bimestre</option></select></label>
                <Field label="Turma" value={exam.className} onChange={(value) => updateExam("className", value)} placeholder="8º A" />
                <label className="field"><span>Turno</span><select value={exam.shift} onChange={(event) => updateExam("shift", event.target.value)}><option>Manhã</option><option>Tarde</option><option>Noite</option><option>Integral</option></select></label>
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
                        <label className="field"><span>Tipo de questão</span><select value={question.type} onChange={(event) => updateQuestion(question.id, { type: event.target.value as QuestionType })}><option value="discursiva">Discursiva</option><option value="multipla">Múltipla escolha</option></select></label>
                        <Field label="Valor" value={question.points} onChange={(value) => updateQuestion(question.id, { points: value })} placeholder="1,0" />
                      </div>
                      <label className="field"><span>Enunciado</span><textarea rows={3} value={question.prompt} placeholder="Digite o enunciado da questão..." onChange={(event) => updateQuestion(question.id, { prompt: event.target.value })} /></label>
                      {question.type === "multipla" && <div className="options-editor"><div className="field-label">Alternativas</div>{question.options.map((option, optionIndex) => <div className="option-row" key={optionIndex}><span>{String.fromCharCode(65 + optionIndex)}</span><input value={option} placeholder={`Alternativa ${String.fromCharCode(65 + optionIndex)}`} onChange={(event) => updateOption(question.id, optionIndex, event.target.value)} /></div>)}</div>}
                    </div>
                  </article>
                ))}
              </div>
              <button className="add-question" onClick={addQuestion} disabled={questions.length >= 10}><Plus size={17} /> Adicionar questão <span>{questions.length >= 10 ? "limite atingido" : `${10 - questions.length} restantes`}</span></button>
            </section>

            <section className="editor-section" id="visual">
              <SectionHeader number="03" eyebrow="ACABAMENTO" title="Um toque final" description="Pequenos detalhes que deixam sua avaliação pronta para imprimir." />
              <Field label="Título da avaliação" value={exam.title} onChange={(value) => updateExam("title", value)} placeholder="Ex.: AVALIAÇÃO DO 3º BIMESTRE" />
              <label className="field"><span>Instruções para a turma</span><textarea rows={2} value={exam.instructions} onChange={(event) => updateExam("instructions", event.target.value)} placeholder="Ex.: Leia cada questão com atenção..." /></label>
              <div className="tip-card"><Sparkles size={18} /><div><strong>Feita para caber melhor</strong><p>A prévia usa duas colunas depois do título para economizar papel sem perder espaço para respostas.</p></div></div>
              <button className="reset-button" onClick={resetDraft}><RotateCcw size={15} /> Restaurar modelo inicial</button>
            </section>
          </div>
          <div className="editor-footer"><button className="save-button" onClick={saveDraft}><Save size={16} /> Salvar rascunho</button><span>As alterações ficam neste dispositivo.</span></div>
        </aside>

        <section className="preview-panel">
          <div className="preview-toolbar">
            <div><div className="eyebrow">PRÉVIA DA FOLHA</div><h2>Veja antes de baixar</h2></div>
            <div className="preview-actions"><button className="secondary-button" onClick={exportPdf}><Printer size={16} /> Exportar PDF</button><button className="primary-button" onClick={exportDocx}><FileDown size={16} /> Baixar DOCX</button></div>
          </div>
          <div className="preview-stage">
            <div className="paper" id="paper-preview">
              <div className="paper-header-table">
                <div className="paper-logo-cell">{logoPreview ? <img src={logoPreview} alt="Logo" /> : <div className="logo-placeholder">LOGO<br />DA ESCOLA</div>}</div>
                <div className="paper-school-cell">{exam.schoolName || "Nome da escola"} – {exam.level || "Nível de ensino"}</div>
                <div className="paper-info-row"><span><b>DISCIPLINA:</b> {exam.subject || "________________"}</span><span><b>PROFESSOR(A):</b> {exam.teacher || "________________"}</span><span><b>BIMESTRE:</b> {exam.bimester || "____"}</span></div>
                <div className="paper-info-row"><span><b>ALUNO(A):</b> {exam.student || ""}</span><span><b>ANO:</b> {exam.className || "____"}</span><span><b>TURMA:</b> __________</span><span><b>TURNO:</b> {exam.shift || "____"}</span><span><b>DATA:</b> {exam.dateDay || "____"}/{exam.dateMonth || "____"}/2026</span></div>
              </div>
              <div className="paper-title-block"><h1>{exam.title || "Título da avaliação"}</h1><p>{exam.instructions || "Leia cada questão com atenção e responda com clareza."}</p></div>
              <div className="student-fields"><div><span>Aluno(a)</span><strong>{exam.student || ""}</strong></div><div><span>Professor(a)</span><strong>{exam.teacher || ""}</strong></div><div className="small-field"><span>Data</span><strong>{exam.dateDay || "____"} / {exam.dateMonth || "____"}</strong></div><div className="small-field"><span>Turno</span><strong>{exam.shift || "____"}</strong></div></div>
              <div className="paper-columns">{questions.map((question, index) => <QuestionPreview question={question} index={index} key={question.id} />)}</div>
              <div className="paper-footer"><span>prova pronta</span><span>{questions.length} questão{questions.length === 1 ? "" : "ões"} <b>·</b> boa prova!</span></div>
            </div>
          </div>
          <div className="preview-note"><Check size={15} /><span>Formato A4 · margens seguras · divisão em duas colunas após o título</span><ChevronDown size={15} /></div>
        </section>
      </main>
    </div>
  );
}
