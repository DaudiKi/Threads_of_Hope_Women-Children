const docx = require('docx');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageNumber,
        Header, Footer, LevelFormat, convertInchesToTwip } = docx;
const fs = require('fs');

const BODY = "Times New Roman", HEAD = "Arial";
const DBL = { line: 480, before: 0, after: 0 };          // double spacing
const IND = { firstLine: convertInchesToTwip(0.5) };

// ---------- helpers ----------
const t = (text, o = {}) => new TextRun({ text, font: BODY, size: 24, ...o });
const ti = (text) => t(text, { italics: true });

// body paragraph, first-line indented, double spaced, justified-left (APA = flush left)
const p = (...runs) => new Paragraph({
  children: runs.map(r => typeof r === "string" ? t(r) : r),
  spacing: DBL, indent: IND, alignment: AlignmentType.LEFT,
});

// body paragraph with no first-line indent
const pn = (...runs) => new Paragraph({
  children: runs.map(r => typeof r === "string" ? t(r) : r),
  spacing: DBL, alignment: AlignmentType.LEFT,
});

// APA Level 1 heading: centered, bold
const h1 = (text) => new Paragraph({
  children: [new TextRun({ text, font: HEAD, size: 28, bold: true })],
  spacing: DBL, alignment: AlignmentType.CENTER, heading: HeadingLevel.HEADING_1,
});
// APA Level 2: flush left, bold
const h2 = (text) => new Paragraph({
  children: [new TextRun({ text, font: HEAD, size: 28, bold: true })],
  spacing: DBL, alignment: AlignmentType.LEFT, heading: HeadingLevel.HEADING_2,
});
// APA Level 3: flush left, bold italic
const h3 = (text) => new Paragraph({
  children: [new TextRun({ text, font: HEAD, size: 24, bold: true, italics: true })],
  spacing: DBL, alignment: AlignmentType.LEFT, heading: HeadingLevel.HEADING_3,
});

// numbered/lettered list item, hanging
const li = (label, ...runs) => new Paragraph({
  children: [t(label + "\u0009"), ...runs.map(r => typeof r === "string" ? t(r) : r)],
  spacing: DBL,
  indent: { left: convertInchesToTwip(1.0), hanging: convertInchesToTwip(0.5) },
});

// reference entry: hanging indent, double spaced
const ref = (...runs) => new Paragraph({
  children: runs.map(r => typeof r === "string" ? t(r) : r),
  spacing: DBL,
  indent: { left: convertInchesToTwip(0.5), hanging: convertInchesToTwip(0.5) },
});

const blank = () => new Paragraph({ children: [t("")], spacing: DBL });

const TITLE = "Offline-First Creative Assessment in Low-Resource Ugandan Schools: Enhancing Developmental Screening, Arts and Design Skills Progression, and Learner Portfolio Ownership Without Requiring Literacy or Internet Connectivity";

// ---------- TITLE PAGE ----------
const titlePage = [
  blank(), blank(), blank(),
  new Paragraph({ children: [new TextRun({ text: TITLE, font: BODY, size: 24, bold: true })],
    spacing: DBL, alignment: AlignmentType.CENTER }),
  blank(),
  new Paragraph({ children: [t("Karabo Ojiambo")], spacing: DBL, alignment: AlignmentType.CENTER }),
  new Paragraph({ children: [t("African Leadership University")], spacing: DBL, alignment: AlignmentType.CENTER }),
  new Paragraph({ children: [t("BSE: Software Engineering — Full-Stack Web Development")], spacing: DBL, alignment: AlignmentType.CENTER }),
  new Paragraph({ children: [t("Unit Two Assignment: Project Draft")], spacing: DBL, alignment: AlignmentType.CENTER }),
  new Paragraph({ children: [t("[Instructor name]")], spacing: DBL, alignment: AlignmentType.CENTER }),
  new Paragraph({ children: [t("October 2, 2025")], spacing: DBL, alignment: AlignmentType.CENTER,
    pageBreakBefore: false }),
  new Paragraph({ children: [new docx.PageBreak()] }),
];

// ---------- CHAPTER ONE ----------
const body = [
h1("Chapter One: Introduction"),

h2("1.1 Introduction and Background"),

p("The use of digital technology to support teaching, learning, and assessment has expanded rapidly across education systems worldwide. In low-resource settings, however, the tools that reach classrooms are seldom designed for the conditions those classrooms actually face. This study examines how an offline-first web application can use children’s drawings and spoken narratives—forms of expression that require neither literacy nor an internet connection—to support developmental screening and structured creative skills progression among vulnerable children aged 8 to 18 in community and faith-based schools in Kampala, Uganda. The section that follows provides the background to this investigation."),

p("Technological change has reshaped how educational institutions deliver instruction, manage records, and assess learning. Over the past two decades, digital systems and data-driven platforms have been adopted across both public and private sectors to improve efficiency and widen access. Within education, this trend has given rise to Educational Technology (EdTech), which refers to the use of digital tools and software to support instructional and administrative processes. Examples include learning management systems, digital assessment platforms, and school information systems. EdTech has attracted sustained attention for its potential to improve learning outcomes, widen participation, and support evidence-based decision-making by teachers and administrators (Means et al., 2013). More recent analysis has cautioned, however, that the benefits of educational technology are unevenly distributed and depend heavily on whether a given tool matches the infrastructural and pedagogical realities of the setting into which it is introduced (UNESCO, 2023)."),

p("Nowhere is that mismatch more consequential than in Sub-Saharan Africa, where education systems have expanded access substantially while learning outcomes have lagged behind. Enrolment growth over recent decades has not translated into proportional gains in foundational skills. An estimated 89% of children in Sub-Saharan Africa are unable to read and understand a simple age-appropriate text by the age of ten, the highest rate of learning poverty recorded in any world region (World Bank et al., 2022). The gap between participation and attainment indicates that the binding constraint in these systems is no longer access to schooling alone, but the quality and responsiveness of what takes place inside the classroom."),

p("Uganda illustrates this pattern clearly. Since the introduction of Universal Primary Education in 1997, enrolment has risen sharply, but the expansion has placed sustained pressure on teaching capacity, classroom infrastructure, and assessment practice. The national pupil-to-teacher ratio in primary education stands at approximately 43:1 (UNESCO Institute for Statistics, 2023), above the 40:1 ratio the United Nations recommends, and individual attention is scarcest in the urban informal settlements surrounding Kampala (Uganda Ministry of Education and Sports, 2022). A substantial share of provision in these areas is delivered by community and faith-based schools and residential child-care institutions, which absorb children the public system does not reach, including orphans, children separated from their families, and children placed in institutional care. These organisations typically operate with minimal budgets, few or no specialist staff, and intermittent electricity and connectivity."),

p("Educational failure in such settings frequently has developmental antecedents that precede formal schooling. Global evidence indicates that approximately 250 million children under the age of five in low- and middle-income countries, representing around 43% of that age group, are at risk of not reaching their developmental potential (Black et al., 2017). Child development proceeds across motor, cognitive, language, and socio-emotional domains, and it advances unevenly rather than in uniform steps. A child may acquire expressive language on schedule while fine-motor control lags, or present age-appropriate motor skills alongside a narrowing of socio-emotional engagement that reflects circumstances at home. Identifying such divergence early is what allows intervention to occur while it is still inexpensive and effective."),

p("Developmental screening is the practice of brief, routine assessment intended to identify children who may benefit from fuller evaluation. Validated instruments for this purpose exist and are well established. The Ages and Stages Questionnaires provide a structured parent-completed monitoring system across developmental domains (Squires & Bricker, 2009), while the Malawi Developmental Assessment Tool was developed and validated specifically for use in African settings, addressing the cultural and linguistic limitations of instruments normed elsewhere (Gladstone et al., 2010). These instruments are psychometrically sound. They are also administratively demanding, typically requiring twenty to thirty minutes of one-to-one administration by a trained assessor, followed by manual scoring and record-keeping."),

p("The practical consequence in high-ratio, low-resource classrooms is that screening does not occur at all. Teachers rely instead on attendance registers and periodic written examinations, neither of which was designed to detect developmental divergence. Written examination carries a further structural problem in this context: it presumes literacy, which is precisely the capability most likely to be compromised in a child with an undetected delay. The dominant assessment instrument is therefore least able to identify the children who most need identifying. Approximately one in ten children worldwide, close to 240 million, lives with a disability, yet under-identification in low-income settings is widely documented and affected children are frequently recognised only after sustained academic failure has already occurred (UNICEF, 2021)."),

p("Research in developmental psychology has long recognised children’s creative output as a window onto capability that verbal and written testing cannot reach. Drawing of the human figure has been used as an index of cognitive maturity since the early twentieth century, formalised by Harris (1963) and subsequently standardised into a quantitative scoring system by Naglieri (1988). Oral narrative offers a parallel measure of language development, in which utterance length, lexical diversity, and the presence of causal and temporal connectives serve as established markers of linguistic and cognitive progression. The pedagogical tradition associated with Reggio Emilia extends this principle into practice, treating drawing, modelling, movement, and spoken story as legitimate languages through which children express understanding rather than as enrichment peripheral to real learning (Edwards et al., 2012). A substantial review of arts education conducted for the Organisation for Economic Co-operation and Development found associations between arts instruction and gains in motor, spatial, and verbal capability, while noting that evidence quality varies by domain (Winner et al., 2013)."),

p("Assessment built on creative production carries pedagogical advantages beyond its diagnostic reach. It is formative rather than punitive, it does not require the child to recognise that assessment is taking place, and it removes the literacy precondition entirely. These properties align with established principles of child-centred practice. Vygotsky (1978) established that learning advances when tasks are pitched within the zone of proximal development and supported by appropriate scaffolding, which implies calibration to the individual child rather than to the age cohort. Scaffolding further presupposes a more capable other, which positions mentorship—by trained adults and by older learners who have themselves progressed through the programme—as structural to the pedagogy rather than supplementary to it, and which gives the mission statement’s commitment to holistic mentorship an operational form. Carr and Lee (2012) demonstrate that narrative, portfolio-based assessment produces a richer and more actionable record of development than periodic grading. Culturally sustaining pedagogy requires that curriculum content draw on learners’ own linguistic and cultural resources rather than importing material from elsewhere (Paris & Alim, 2017), and self-determination theory identifies autonomy, competence, and relatedness as the conditions under which sustained engagement occurs (Deci & Ryan, 2000). Universal Design for Learning provides a framework for ensuring that no step in a learning activity depends on a single mode of representation (CAST, 2018). These commitments are reinforced by the child’s right to express views, to seek and impart information, and to participate in play and artistic activity (United Nations, 1989), and by the proposition that children’s participation should progress from consultation toward genuine co-authorship (Hart, 1992)."),

p("Any technological response must contend with the infrastructure actually present rather than the infrastructure assumed. Just 38% of Africa’s population used the internet in 2024, against a global average of 68%, and the continent records the widest urban–rural divide in the world, at 57% in cities against 23% in rural areas (International Telecommunication Union, 2024). Access is predominantly mobile, intermittent, and metered rather than fixed and continuous Mobile is the primary and often the only route to connectivity, with smartphone adoption rising but affordability and data cost remaining material constraints (GSMA, 2024). Educational software that assumes a continuous connection, a per-learner device, and a recurring subscription is structurally unsuited to a community school operating on intermittent power with a small number of shared handsets."),

p("Developments in web platform engineering make an alternative viable. Progressive Web Applications use service workers, client-side storage, and background synchronisation to deliver full application functionality without a continuous network connection, and install to a device directly from the browser without an application store or distributable installer (Biørn-Hansen et al., 2017). In parallel, machine learning inference has become practical within the browser itself, allowing models to execute locally on the user’s device rather than transmitting data to a remote server (Smilkov et al., 2019). This matters twice over in the present context: it removes the connectivity dependency, and it keeps sensitive information about children on the device rather than in transit. Speech processing for the languages concerned is also better resourced than is commonly assumed. A publicly available Luganda radio speech corpus of 155 hours, the first such dataset released in Sub-Saharan Africa, now supports automatic speech recognition work in the language (Mukiibi et al., 2022), alongside open parallel text and speech resources for Luganda and neighbouring Ugandan languages (Sunbird AI, 2024)."),

p("Research and software development involving vulnerable children carry obligations that shape system design rather than merely accompanying it. Processing of personal data in Uganda is governed by the Data Protection and Privacy Act, which establishes requirements for lawful processing, data minimisation, and the rights of data subjects (Republic of Uganda, 2019). Violence against children in Uganda is common and frequently undisclosed, with national survey evidence indicating that three in four young adults experienced some form of violence during childhood and one in three experienced at least two forms of it (Ministry of Gender, Labour and Social Development, 2018). Any system capable of surfacing indicators of distress is therefore a safeguarding instrument as well as an educational one, and must be governed accordingly, with human confirmation of every indicator and no automated escalation."),

p("Taken together, these conditions describe a specific and currently unmet need. Community and faith-based schools in Kampala serve children at elevated developmental risk, lack the staff and instruments to screen them, and operate on infrastructure that excludes the software products designed for better-resourced systems. At the same time, these children produce creative work—drawings and spoken stories—that developmental research establishes as a valid source of signal, and the web platform has matured to the point where that signal can be extracted locally, offline, on the devices these schools already possess. This study addresses the integration of these elements into a single system, provisionally named Loom, that screens, teaches, and accumulates a record the child ultimately owns: a cumulative creative archive referred to throughout as the Golden Thread portfolio."),

p("This study is positioned within the mission statement guiding the wider project: to transform foundational education for vulnerable children by partnering with community and faith-based schools to integrate creative arts, design-based learning, and holistic mentorship, thereby empowering the next generation to break cycles of marginalisation. It addresses Education, one of the fourteen Global Challenges, by targeting the undetected early learning loss that precedes dropout rather than the dropout itself. It pursues that aim through the Global Opportunity of Arts, Design and Culture, treating creative drawing, visual design, and oral storytelling as the primary diagnostic and pedagogical medium for children whose literacy cannot be assumed."),

// ---------- 1.2 ----------
h2("1.2 Problem Statement"),

p("Grassroots community and faith-based schools and child-care institutions in Kampala lack any objective instrument for tracking the non-linear development of the children in their care. Overworked educators manage large classes using manual paper registers and rote written testing, with the result that fine-motor impairment, cognitive delay, language difficulty, and the behavioural signatures of domestic trauma commonly go undetected until a child has already failed repeatedly and, in many cases, left school. Two categories of existing solution come closest to addressing this problem, and each falls short in ways that define the gap this project addresses."),

p("The first is commercial school management and analytics software. Platforms such as Zeraki Analytics, which is widely deployed across East African schools, and QuickSchools provide attendance tracking, examination grade capture, report-card generation, and performance dashboards (Zeraki, 2024; QuickSchools, 2024). These products are mature and effective within their intended scope, but that scope is secondary-school examination administration. They are cloud-dependent and subscription-priced, which places them beyond both the connectivity and the budget of the institutions described above. More fundamentally, they record only binary attendance and numerical examination marks. They capture outcome rather than development, they offer no representation whatsoever for a learner who cannot yet read, and they contain no construct corresponding to motor control, spatial reasoning, narrative capability, or socio-emotional change."),

p("The second is standardised paper-based developmental and psychometric instrumentation, including the Ages and Stages Questionnaires (Squires & Bricker, 2009), the Draw-A-Person quantitative scoring system (Naglieri, 1988), and the Malawi Developmental Assessment Tool (Gladstone et al., 2010). These instruments are clinically validated and, in the case of the latter, specifically normed for African populations. Their limitation is operational rather than scientific. Each requires approximately twenty to thirty minutes of individual administration by a trained assessor, followed by manual scoring. At the pupil-to-teacher ratios prevailing in Kampala’s community schools, and with no child psychologist on staff, the arithmetic does not permit their use: screening a single class of forty children once would consume between thirteen and twenty hours of specialist time that the institution does not have. In practice these instruments are not merely difficult to apply in such settings—they are abandoned, and the resulting paper records never become a longitudinal trend."),

p("A third and more recent category, consumer creative and learning applications, merits brief mention. Products in this group support drawing and guided learning activity but are cloud-first, English-first, and designed around a child with a personal device and a parent-managed account. They generate creative output without producing measurement, provide no educator view, and offer no safeguarding pathway or portable record."),

p("The gap is therefore specific. No existing system is simultaneously offline-first, multimodal across visual and oral expression, longitudinal across the 8-to-18 period, and owned by the child rather than the institution. None converts the creative work these classrooms already produce into developmental signal, an adaptive curriculum, a safeguarding early warning, and a portable record of capability, on the low-cost shared devices these schools actually possess. Closing this gap is the direct operational expression of the mission statement\u2019s commitment to transforming foundational education for vulnerable children through creative arts and design-based learning, and it addresses the Education Global Challenge at its root: the undetected developmental divergence that precedes school failure, rather than the failure itself."),

// ---------- 1.3 ----------
h2("1.3 Project’s Main Objective"),

p("To design, develop, and evaluate an offline-first progressive web application that uses children’s drawings, designs, and recorded oral narratives as both a continuous developmental screening instrument and an adaptive arts, design, and culture curriculum for learners aged 8 to 18 in low-resource Kampala schools, addressing the diagnostic, infrastructural, and continuity gaps identified in the problem statement, in order to surface developmental concerns while intervention remains effective and to equip each learner with a verified, exportable portfolio of creative capability by the age of eighteen, in direct service of the mission statement\u2019s commitment to empowering the next generation to break cycles of marginalisation."),

h3("1.3.1 Specific Objectives"),

li("SO1.", "To design and implement a child-safe account and creative studio module supporting canvas drawing, voice recording, and photographic capture of physical artwork, operable without reading, and to validate it with 60 learners aged 8 to 18 across two partner institutions within four months of project commencement."),
li("SO2.", "To develop a tiered on-device multimodal scoring pipeline producing at least six developmental indices per submission, comprising language-independent acoustic features and Luganda narrative-connective keyword spotting executed locally, with full transcription deferred to synchronisation, returning an on-device result in under three seconds on a 2 GB RAM Android device with no network connection, by month six."),
li("SO3.", "To construct a four-stage, four-strand arts, design, and culture competency framework spanning ages 8 to 18, encoded as a machine-readable progression graph, and to validate its stage descriptors with a minimum of six practising educators and two subject specialists prior to field deployment."),
li("SO4.", "To implement an offline-first synchronisation layer using service workers, an IndexedDB operation queue, and idempotent batch replay, capable of supporting a complete 40-minute classroom session with zero connectivity and reconciling without data loss, demonstrated across 100 queued operations per device, by month seven."),
li("SO5.", "To deliver an educator console presenting a cohort development heatmap and a human-confirmed early-warning triage queue, and to evaluate it across a 12-week field deployment against educator-reported usefulness and measured time-to-review per child."),
li("SO6.", "To produce and evaluate a learner-owned portfolio export comprising a creative archive, skills record, and revocable public link, with a minimum of 15 learners in the oldest age band, assessed against the published entry requirements of local art schools, apprenticeship programmes, and creative employers, by month ten."),
li("SO7.", "To implement and evaluate a structured mentorship module pairing each learner with a trained adult mentor drawn from the existing community cooperative and with an older peer who has progressed through the programme, recording mentoring hours as an assessed competency, and to measure its association with sustained engagement across the 12-week field deployment."),

// ---------- 1.4 ----------
h2("1.4 Research Questions"),

li("RQ1.", "To what extent can children aged 8 to 18 with limited or no literacy complete an unassisted creative submission, comprising a drawing and a recorded spoken narrative, using a browser-based studio on a shared low-cost device?"),
li("RQ2.", "How accurately do on-device visual features and language-independent acoustic features extracted from children’s drawings and oral narratives correspond to developmental indicators assessed by a trained practitioner using established instruments?"),
li("RQ3.", "Which four-stage arts, design, and culture competency progression do practising educators and subject specialists validate as both developmentally appropriate and economically relevant for learners aged 8 to 18 in this setting?"),
li("RQ4.", "What offline-first architecture permits a complete classroom session to be conducted and subsequently reconciled without data loss under the connectivity and device constraints of low-resource Kampala schools?"),
li("RQ5.", "How does continuous creative-work screening alter the time educators spend identifying developmental concerns, and how early are concerns surfaced relative to current practice?"),
li("RQ6.", "What portfolio artefact, accumulated across the 8-to-18 period, do local art schools, apprenticeship programmes, and creative employers recognise as credible evidence of capability?"),
li("RQ7.", "Can closed-vocabulary keyword spotting over Luganda narrative connectives recover narrative-structure signal from children’s spoken stories without open-vocabulary speech recognition, and how does its output compare with transcription by a fluent adult listener?"),

// ---------- 1.5 ----------
li("RQ8.", "What association does structured mentorship, pairing learners with adult mentors from the community cooperative and with older peers, have with sustained engagement and completion of creative submissions across the deployment period?"),

h2("1.5 Project Scope"),

p("The project is bounded geographically, demographically, temporally, and technically in order to remain deliverable within a single development cycle."),

p("Geographically, deployment is limited to two partner institutions in Kampala, Uganda: one community and faith-based preparatory school and one residential child-care centre, both operating at low resource levels with intermittent connectivity and shared rather than per-learner devices. The institutions are described by characteristic rather than named, since naming a small care institution alongside a described population of vulnerable children would narrow identification unacceptably; sites will be named only once letters of access are secured and the institutions have approved the mention. National rollout, deployment elsewhere in Uganda, and any cross-border work are outside scope."),

p("Demographically, the study population comprises 60 to 120 learners aged 8 to 18 distributed across four developmental stage bands, together with six educators and administrators and eight women drawn from the existing community cooperative who will act as creative mentors. Children under the age of eight, out-of-school youth, and learners in well-resourced private schools are outside scope."),

p("Temporally, the work follows a ten-month development cycle incorporating a twelve-week supervised field deployment. Multi-year longitudinal validation of the developmental indices, which the design ultimately implies, is explicitly named as future work and is not attempted here."),

p("Technically, the artefact is an offline-first progressive web application built on React with Vite and TypeScript, a Node.js and Express backend using Sequelize, and PostgreSQL with object storage, extending an existing platform codebase. Inference is tiered and executed on-device. The interface is delivered in English, Kiswahili, French, Luganda, and Kinyarwanda, each with recorded audio instruction so that no step depends on reading. Speech analysis is scoped to Luganda alone; the remaining four languages are interface locales only. Native mobile builds, real-time video, blockchain credentialing, full open-vocabulary on-device transcription, and automated clinical diagnosis are all outside scope. The wider platform’s production and revenue pathway, under which learner designs are manufactured by the women’s cooperative and sold to fund learner education accounts, is deliberately excluded from this cycle: it raises child-labour, consent, and financial-safeguarding questions that warrant separate ethical review, and it is named here as future work rather than attempted alongside the screening deployment."),

p("In terms of hardware, the target is Android 9 or later running Chrome on devices with 2 GB of RAM, together with staffroom Windows laptops, at a ratio of approximately one shared device per five to eight learners. Procurement of devices for individual learners is not contemplated. With respect to assessment, the system performs screening and early warning only; all output is advisory and subject to educator confirmation, and no automated referral occurs without a human decision. Ethically, the work is bounded by institutional ethics approval, guardian consent and child assent for every participant, data minimisation, local-first storage, and compliance with the Data Protection and Privacy Act (Republic of Uganda, 2019)."),

// ---------- 1.6 ----------
h2("1.6 Significance and Justification"),

p("Upon successful implementation, educators and their pupils would gain developmental screening where none currently exists. A teacher responsible for forty would see a motor delay or a widening withdrawal within weeks, not at repeated grade failure."),

p("For learners and guardians, ten years of creative work would compound rather than evaporate, equipping each young person at eighteen with the portfolio, specialism, and mentor reference needed to break the cycle of marginalisation the mission statement targets (Winner et al., 2013)."),

p("For partner institutions, policymakers, and the wider sector, it would show multimodal screening running offline on hardware these schools already own, removing the subscription, connection, and specialist that make such screening a privilege and the Education Global Challenge persistent."),

// ---------- REFERENCES ----------
new Paragraph({ children: [new docx.PageBreak()] }),
h1("References"),

ref("Biørn-Hansen, A., Majchrzak, T. A., & Grønli, T.-M. (2017). Progressive web apps: The possible web-native unifier for mobile development. In ", ti("Proceedings of the 13th International Conference on Web Information Systems and Technologies"), " (pp. 344–351). SCITEPRESS. https://doi.org/10.5220/0006353703440351"),

ref("Black, M. M., Walker, S. P., Fernald, L. C. H., Andersen, C. T., DiGirolamo, A. M., Lu, C., McCoy, D. C., Fink, G., Shawar, Y. R., Shiffman, J., Devercelli, A. E., Wodon, Q. T., Vargas-Barón, E., & Grantham-McGregor, S. (2017). Early childhood development coming of age: Science through the life course. ", ti("The Lancet, 389"), "(10064), 77–90. https://doi.org/10.1016/S0140-6736(16)31389-7"),

ref("Carr, M., & Lee, W. (2012). ", ti("Learning stories: Narrative assessment in early childhood education"), ". SAGE Publications."),

ref("CAST. (2018). ", ti("Universal design for learning guidelines version 2.2"), ". https://udlguidelines.cast.org"),

ref("Deci, E. L., & Ryan, R. M. (2000). The “what” and “why” of goal pursuits: Human needs and the self-determination of behavior. ", ti("Psychological Inquiry, 11"), "(4), 227–268. https://doi.org/10.1207/S15327965PLI1104_01"),

ref("Edwards, C., Gandini, L., & Forman, G. (Eds.). (2012). ", ti("The hundred languages of children: The Reggio Emilia experience in transformation"), " (3rd ed.). Praeger."),

ref("Gladstone, M., Lancaster, G. A., Umar, E., Nyirenda, M., Kayira, E., van den Broek, N. R., & Smyth, R. L. (2010). The Malawi Developmental Assessment Tool (MDAT): The creation, validation, and reliability of a tool to assess child development in rural African settings. ", ti("PLoS Medicine, 7"), "(5), e1000273. https://doi.org/10.1371/journal.pmed.1000273"),

ref("GSMA. (2024). ", ti("The mobile economy Sub-Saharan Africa 2024"), ". GSMA Intelligence. https://www.gsma.com/solutions-and-impact/connectivity-for-good/mobile-economy/"),

ref("Harris, D. B. (1963). ", ti("Children’s drawings as measures of intellectual maturity: A revision and extension of the Goodenough Draw-a-Man Test"), ". Harcourt, Brace & World."),

ref("Hart, R. A. (1992). ", ti("Children’s participation: From tokenism to citizenship"), " (Innocenti Essays No. 4). UNICEF International Child Development Centre."),

ref("International Telecommunication Union. (2024). ", ti("Measuring digital development: Facts and figures 2024"), ". ITU Publications. https://www.itu.int/itu-d/reports/statistics/"),

ref("Means, B., Toyama, Y., Murphy, R., & Bakia, M. (2013). The effectiveness of online and blended learning: A meta-analysis of the empirical literature. ", ti("Teachers College Record, 115"), "(3), 1–47."),

ref("Ministry of Gender, Labour and Social Development. (2018). ", ti("Uganda violence against children survey: Findings from a national survey 2015"), ". Government of Uganda."),

ref("Mukiibi, J., Katumba, A., Nakatumba-Nabende, J., Hussein, A., & Meyer, J. (2022). The Makerere radio speech corpus: A Luganda radio corpus for automatic speech recognition. In ", ti("Proceedings of the Thirteenth Language Resources and Evaluation Conference"), " (pp. 1945–1954). European Language Resources Association. https://aclanthology.org/2022.lrec-1.208"),

ref("Naglieri, J. A. (1988). ", ti("Draw A Person: A quantitative scoring system"), ". The Psychological Corporation."),

ref("Paris, D., & Alim, H. S. (Eds.). (2017). ", ti("Culturally sustaining pedagogies: Teaching and learning for justice in a changing world"), ". Teachers College Press."),

ref("QuickSchools. (2024). ", ti("QuickSchools school management system"), ". https://www.quickschools.com"),

ref("Republic of Uganda. (2019). ", ti("The Data Protection and Privacy Act, 2019"), ". Uganda Printing and Publishing Corporation."),

ref("Smilkov, D., Thorat, N., Assogba, Y., Yuan, A., Kreeger, N., Yu, P., Zhang, K., Cai, S., Nielsen, E., Soergel, D., Bileschi, S., Terry, M., Nicholson, C., Gupta, S. N., Sirajuddin, S., Sculley, D., Monga, R., Corrado, G., Viégas, F. B., & Wattenberg, M. (2019). TensorFlow.js: Machine learning for the web and beyond. In ", ti("Proceedings of the 2nd SysML Conference"), ". Palo Alto, CA."),

ref("Squires, J., & Bricker, D. (2009). ", ti("Ages & Stages Questionnaires (ASQ-3): A parent-completed child monitoring system"), " (3rd ed.). Paul H. Brookes Publishing."),

ref("Sunbird AI. (2024). ", ti("SALT: Sunbird African Language Technology dataset and models"), ". https://github.com/SunbirdAI/salt"),

ref("Uganda Ministry of Education and Sports. (2022). ", ti("Education and sports sector statistical abstract"), ". Government of Uganda."),

ref("UNESCO. (2023). ", ti("Global education monitoring report 2023: Technology in education — A tool on whose terms?"), " UNESCO Publishing. https://doi.org/10.54676/UZQV8501"),

ref("UNESCO Institute for Statistics. (2023). ", ti("Education indicators: Pupil-teacher ratio, primary"), ". UIS.Stat. https://data.uis.unesco.org"),

ref("UNICEF. (2021). ", ti("Seen, counted, included: Using data to shed light on the well-being of children with disabilities"), ". United Nations Children’s Fund. https://data.unicef.org/resources/children-with-disabilities-report-2021/"),

ref("United Nations. (1989). ", ti("Convention on the Rights of the Child"), ". United Nations Treaty Series, 1577, 3."),

ref("Vygotsky, L. S. (1978). ", ti("Mind in society: The development of higher psychological processes"), ". Harvard University Press."),

ref("Winner, E., Goldstein, T. R., & Vincent-Lancrin, S. (2013). ", ti("Art for art’s sake? The impact of arts education"), ". OECD Publishing. https://doi.org/10.1787/9789264180789-en"),

ref("World Bank, UNESCO, UNICEF, Foreign, Commonwealth & Development Office, United States Agency for International Development, & Bill & Melinda Gates Foundation. (2022). ", ti("The state of global learning poverty: 2022 update"), ". World Bank."),

ref("Zeraki. (2024). ", ti("Zeraki Analytics: School analytics and management"), ". https://zeraki.app"),
];

const doc = new Document({
  creator: "Karabo Ojiambo",
  lastModifiedBy: "Karabo Ojiambo",
  title: "Unit Two Assignment: Project Draft",
  description: "Chapter One of a research proposal.",
  styles: { default: { document: { run: { font: BODY, size: 24 } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    headers: {
      default: new Header({ children: [ new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [ new TextRun({ children: [PageNumber.CURRENT], font: BODY, size: 24 }) ],
      })]}),
    },
    children: [...titlePage, ...body],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync("Ojiambo_Unit Two Assignment.docx", b);
  console.log("written:", b.length, "bytes");
});
