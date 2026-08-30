/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package timetable;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.Collections;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

/**
 * Формирует недельное расписание богослужений на основе календаря
 * azbyka.ru и шаблона {@code template.docx}.
 *
 * <p>По сравнению с исходной версией здесь сделано следующее:</p>
 * <ol>
 *   <li><b>Лишние пробелы перед знаками препинания.</b> Раньше текст с сайта
 *       вставлялся как есть, из-за чего иногда получалось "Пятидесятнице .".
 *       Теперь весь текст, полученный с сайта, проходит через {@link #clean},
 *       которая убирает пробелы перед .,!?;: и схлопывает повторные пробелы.
 *       Это универсальное решение — оно не зависит от того, из-за какого
 *       именно тега на странице появился лишний пробел.</li>
 *   <li><b>Дата в три строки.</b> Раньше "31 августа" выводилось одной
 *       строкой (либо переносилось само, "на глаз", если не помещалось).
 *       Теперь {@link #insertDate} явно вставляет разрыв строки между
 *       числом и названием месяца, а также — если в шаблоне день недели
 *       идёт в том же абзаце через пробел (как у понедельника) — превращает
 *       этот пробел в ещё один разрыв строки. В результате всегда получается
 *       "31" / "августа" / "понедельник" на трёх строках.</li>
 *   <li><b>Автоподбор шрифта под одну страницу.</b> После сборки документа
 *       {@link PageFitter} рендерит его в PDF (через LibreOffice) и, пока
 *       страниц больше одной, уменьшает шрифт на 1 пункт — по очереди,
 *       начиная с ячейки "Чтим память" с самым длинным текстом, затем со
 *       следующей по длине и так по кругу. Первая жирная фраза дня
 *       ("Седмица ... . Поста нет. Глас N-й.", т.е. div.post + первый
 *       &lt;p&gt; из div.text.day__text) в уменьшении шрифта не участвует.</li>
 * </ol>
 *
 * <p><b>Зависимости:</b> jsoup, poi + poi-ooxml (как и раньше).
 * Для пункта 3 дополнительно нужны установленные и доступные в PATH
 * LibreOffice ({@code soffice}) и Poppler ({@code pdfinfo}) — ими документ
 * конвертируется в PDF и считаются страницы. Если этих утилит нет,
 * расписание всё равно будет сохранено — просто без автоподбора шрифта,
 * см. {@link PageFitter#fit}.</p>
 */
public final class Timetable {

    // ---- Настройки -----------------------------------------------------

    private static final String TEMPLATE_FILE = "template.docx";
    private static final String AZBYKA_DAY_URL_PREFIX = "https://azbyka.ru/days/";

    /** Порядок плейсхолдеров даты в шаблоне: пн..вс. */
    private static final String[] DATE_PLACEHOLDERS = {"mon", "tue", "wen", "thu", "fri", "sat", "sun"};
    /** Порядок плейсхолдеров текста дня в столбце "Чтим память": пн..вс. */
    private static final String[] TEXT_PLACEHOLDERS = {"one", "two", "three", "four", "five", "six", "seven"};

    /** Родительный падеж месяцев — азбука.ру и Word одинаково ждут "31 августа", а не "31 август". */
    private static final String[] MONTHS_GENITIVE = {
            "января", "февраля", "марта", "апреля", "мая", "июня",
            "июля", "августа", "сентября", "октября", "ноября", "декабря"
    };

    /** Размер жирной "шапки" дня (Седмица .../Глас N-й) — фиксирован и не участвует в автоподборе. */
    private static final int HEADER_FONT_SIZE_PT = 11;
    /** Базовый размер обычного текста в столбце "Чтим память" (берётся из template.docx). */
    private static final int BODY_FONT_SIZE_PT = 10;
    /** Нижняя граница, ниже которой автоподбор шрифт уже не уменьшает. */
    private static final int MIN_FONT_SIZE_PT = 6;
    /** Предохранитель от бесконечного цикла в автоподборе. */
    private static final int MAX_SHRINK_STEPS = 80;

    private static final Pattern WHITESPACE = Pattern.compile("\\s+");
    private static final Pattern SPACE_BEFORE_PUNCTUATION = Pattern.compile("\\s+([.,!?;:])");

    private Timetable() {
        // утилитный класс, точка входа — main()
    }

    // ---- Точка входа ------------------------------------------------------

    public static void main(String[] args) throws Exception {
        LocalDate monday = LocalDate.now().with(TemporalAdjusters.nextOrSame(DayOfWeek.MONDAY));

        List<DaySchedule> week = new ArrayList<>(DATE_PLACEHOLDERS.length);
        for (int i = 0; i < DATE_PLACEHOLDERS.length; i++) {
            LocalDate date = monday.plusDays(i);
            System.out.println("Загружаю " + date + " ...");
            week.add(fetchDaySchedule(date));
        }

        XWPFDocument document = new XWPFDocument(OPCPackage.open(TEMPLATE_FILE));
        try {
            List<MemoryCell> memoryCells = new ArrayList<>(week.size());
            for (int i = 0; i < week.size(); i++) {
                DaySchedule day = week.get(i);
                insertDate(document, DATE_PLACEHOLDERS[i], day.getDate());
                memoryCells.add(fillMemoryCell(document, TEXT_PLACEHOLDERS[i], day));
            }

            Path outputPath = Paths.get(buildOutputFileName(monday));
            saveDocument(document, outputPath);

            PageFitter.fit(document, outputPath, memoryCells);

            System.out.println("Готово: " + outputPath.toAbsolutePath());
        } finally {
            document.close();
        }
    }

    private static String buildOutputFileName(LocalDate monday) {
        return "Расписание на " + monday.getDayOfMonth() + " " + MONTHS_GENITIVE[monday.getMonthValue() - 1]
                + " " + monday.getYear() + ".docx";
    }

    private static void saveDocument(XWPFDocument document, Path path) throws IOException {
        try (OutputStream out = Files.newOutputStream(path)) {
            document.write(out);
        }
    }

    // ---- Получение данных с azbyka.ru -------------------------------------

    /** Данные одного дня: дата, жирная "шапка" и список фраз для выделения. */
    private static final class DaySchedule {

        private final LocalDate date;
        private final String header;
        private final String fullText;
        private final List<String> boldPhrases;
        private final List<String> italicPhrases;

        DaySchedule(LocalDate date, String header, String fullText,
                    List<String> boldPhrases, List<String> italicPhrases) {
            this.date = date;
            this.header = header;
            this.fullText = fullText;
            this.boldPhrases = boldPhrases;
            this.italicPhrases = italicPhrases;
        }

        LocalDate getDate() {
            return date;
        }

        /** "Седмица 14-я по Пятидесятнице. Поста нет. Глас 4-й." — без лишних пробелов перед точками. */
        String getHeader() {
            return header;
        }

        /** Шапка + остальной текст дня (жития, памяти святых и т.д.). */
        String getFullText() {
            return fullText;
        }

        List<String> getBoldPhrases() {
            return boldPhrases;
        }

        List<String> getItalicPhrases() {
            return italicPhrases;
        }
    }

    private static DaySchedule fetchDaySchedule(LocalDate date) throws IOException {
        String url = AZBYKA_DAY_URL_PREFIX + date; // LocalDate.toString() уже даёт yyyy-MM-dd
        Document page;
        try {
            page = Jsoup.connect(url)
                    .userAgent("Mozilla/5.0 (compatible; TimetableBot/1.0)")
                    .timeout(20_000)
                    .get();
        } catch (IOException e) {
            throw new IOException("Не удалось загрузить " + url, e);
        }

        Elements dayText = page.select("div.text.day__text");
        Elements paragraphs = dayText.select("p");
        Elements lists = dayText.select("ul");

        String topText = clean(page.select("div.post").text());
        String firstParagraph = paragraphs.isEmpty() ? "" : clean(paragraphs.get(0).text());
        String header = clean(joinSentences(topText, firstParagraph));

        // Текст дня собираем ДО удаления <em> — иначе фразы, которые ниже
        // выделяются курсивом, попросту пропадут из отображаемого текста.
        String bodyText = clean(lists.text());
        String fullText = clean(joinSentences(header, bodyText));

        List<String> italicPhrases = new ArrayList<>();
        for (Element e : lists.select("li em")) {
            italicPhrases.add(clean(e.text()));
        }
        // Убираем <em> из разметки, чтобы дальше их текст не задвоился
        // при выборке strong/b/span (как и в исходной версии).
        lists.select("li em").remove();

        List<String> boldPhrases = new ArrayList<>();
        for (Element e : lists.select("li strong")) {
            boldPhrases.add(clean(e.text()));
        }
        for (Element e : lists.select("li b")) {
            boldPhrases.add(clean(e.text()));
        }
        for (Element e : lists.select("li span")) {
            String text = clean(e.text());
            if (!text.isEmpty()) {
                boldPhrases.add(text);
            }
        }

        return new DaySchedule(date, header, fullText, boldPhrases, italicPhrases);
    }

    private static String joinSentences(String first, String second) {
        if (first.isEmpty()) {
            return second;
        }
        if (second.isEmpty()) {
            return first;
        }
        return first + ". " + second;
    }

    /** Убирает повторные пробелы/переносы и пробел перед знаками препинания ("Пятидесятнице ." → "Пятидесятнице."). */
    private static String clean(String raw) {
        if (raw == null) {
            return "";
        }
        String result = WHITESPACE.matcher(raw).replaceAll(" ").trim();
        result = SPACE_BEFORE_PUNCTUATION.matcher(result).replaceAll("$1");
        return result;
    }

    // ---- Заполнение шаблона -------------------------------------------------

    /** Ссылка на ячейку "Чтим память" одного дня + runs, которые нельзя трогать при автоподборе шрифта. */
    private static final class MemoryCell {

        private final XWPFTableCell cell;
        private final Set<XWPFRun> protectedRuns;
        private final int approximateLength;

        MemoryCell(XWPFTableCell cell, Set<XWPFRun> protectedRuns, int approximateLength) {
            this.cell = cell;
            this.protectedRuns = protectedRuns;
            this.approximateLength = approximateLength;
        }
    }

    /**
     * Заменяет плейсхолдер даты (mon/tue/...) в ячейке "Число" на дату в
     * формате "31" (разрыв строки) "августа". Если сразу за плейсхолдером в
     * том же абзаце идёт пустой run-разделитель, за которым следует день
     * недели (как в шаблоне для понедельника), этот разделитель тоже
     * превращается в разрыв строки — чтобы день недели оказался на своей
     * отдельной строке.
     */
    private static void insertDate(XWPFDocument document, String placeholder, LocalDate date) {
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (int i = 0; i < runs.size(); i++) {
                            XWPFRun run = runs.get(i);
                            String runText = run.getText(0);
                            if (runText == null || !runText.trim().equals(placeholder)) {
                                continue;
                            }

                            run.setText(String.valueOf(date.getDayOfMonth()), 0);
                            run.addBreak();
                            appendText(run, MONTHS_GENITIVE[date.getMonthValue() - 1]);

                            if (i + 1 < runs.size()) {
                                XWPFRun next = runs.get(i + 1);
                                String nextText = next.getText(0);
                                if (nextText != null && nextText.trim().isEmpty()) {
                                    clearRunText(next);
                                    next.addBreak();
                                }
                            }
                            return; // плейсхолдер в документе один
                        }
                    }
                }
            }
        }
        throw new IllegalStateException("В шаблоне не найден плейсхолдер даты \"" + placeholder + "\"");
    }

    /**
     * Заменяет текстовый плейсхолдер (one/two/...) в ячейке "Чтим память" на
     * текст дня и раскрашивает в нём фразы: bold-фразы жирным, italic-фразы
     * курсивом, а шапку дня — жирным фиксированным размером
     * {@link #HEADER_FONT_SIZE_PT}, защищённым от автоподбора шрифта.
     *
     * @return ссылка на ячейку и на runs, которые нельзя уменьшать при автоподборе
     */
    private static MemoryCell fillMemoryCell(XWPFDocument document, String placeholder, DaySchedule day) {
        XWPFTableCell targetCell = null;
        List<XWPFRun> headerRuns = new ArrayList<>();

        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        boolean matchedInParagraph = false;
                        for (XWPFRun run : paragraph.getRuns()) {
                            String runText = run.getText(0);
                            if (runText != null && runText.contains(placeholder)) {
                                run.setText(runText.replace(placeholder, day.getFullText()), 0);
                                matchedInParagraph = true;
                            }
                        }
                        if (!matchedInParagraph) {
                            continue;
                        }

                        targetCell = cell;
                        for (String phrase : day.getBoldPhrases()) {
                            highlightKeyword(paragraph, phrase, true, false, null);
                        }
                        for (String phrase : day.getItalicPhrases()) {
                            highlightKeyword(paragraph, phrase, false, true, null);
                        }
                        if (!day.getHeader().isEmpty()) {
                            headerRuns.addAll(
                                    highlightKeyword(paragraph, day.getHeader(), true, false, HEADER_FONT_SIZE_PT));
                        }
                    }
                }
            }
        }

        if (targetCell == null) {
            throw new IllegalStateException("В шаблоне не найдена ячейка с плейсхолдером \"" + placeholder + "\"");
        }

        Set<XWPFRun> protectedRuns = Collections.newSetFromMap(new IdentityHashMap<>());
        protectedRuns.addAll(headerRuns);
        return new MemoryCell(targetCell, protectedRuns, day.getFullText().length());
    }

    // ---- Низкоуровневая работа с runs (OOXML) -------------------------------

    /**
     * Ищет во всех runs абзаца вхождения {@code keyword} и выносит их в
     * отдельный run с заданным форматированием, сохраняя исходное
     * форматирование (шрифт и т.п.) остального текста. Аналог формата "найти
     * и подсветить" — если keyword встречается несколько раз, каждое
     * вхождение получает свой run.
     *
     * @return созданные runs с keyword (пустой список, если совпадений не было)
     */
    private static List<XWPFRun> highlightKeyword(XWPFParagraph paragraph, String keyword,
                                                  boolean bold, boolean italic, Integer fixedFontSizePt) {
        List<XWPFRun> created = new ArrayList<>();
        if (keyword == null || keyword.isEmpty()) {
            return created;
        }

        int runIndex = 0;
        while (runIndex < paragraph.getRuns().size()) {
            XWPFRun run = paragraph.getRuns().get(runIndex);
            String runText = run.getText(0);

            if (runText != null && runText.contains(keyword)) {
                StringBuilder buffer = new StringBuilder();
                for (char c : runText.toCharArray()) {
                    buffer.append(c);
                    if (buffer.length() < keyword.length() || !endsWith(buffer, keyword)) {
                        continue;
                    }

                    // текст run'а до найденного keyword остаётся в исходном run
                    run.setText(buffer.substring(0, buffer.length() - keyword.length()), 0);

                    // сам keyword — в новом run с нужным форматированием
                    XWPFRun keywordRun = paragraph.insertNewRun(++runIndex);
                    cloneRunProperties(run, keywordRun);
                    keywordRun.setText(keyword, 0);
                    keywordRun.setBold(bold);
                    keywordRun.setItalic(italic);
                    if (fixedFontSizePt != null) {
                        keywordRun.setFontSize(fixedFontSizePt);
                    }
                    created.add(keywordRun);

                    // новый run для текста, который будет идти после keyword
                    XWPFRun tailRun = paragraph.insertNewRun(++runIndex);
                    cloneRunProperties(run, tailRun);
                    run = tailRun;
                    buffer = new StringBuilder();
                }
                run.setText(buffer.toString(), 0);
            }
            runIndex++;
        }
        return created;
    }

    private static boolean endsWith(StringBuilder buffer, String suffix) {
        return buffer.indexOf(suffix, buffer.length() - suffix.length()) >= 0;
    }

    /** Копирует форматирование (rPr) одного run на другой — как в исходном коде. */
    private static void cloneRunProperties(XWPFRun source, XWPFRun dest) {
        CTR sourceCtr = source.getCTR();
        CTRPr sourceRPr = sourceCtr.getRPr();
        if (sourceRPr != null) {
            dest.getCTR().setRPr((CTRPr) sourceRPr.copy());
        }
    }

    /** Дописывает в run ещё один текстовый узел (после уже существующего текста/разрыва строки). */
    private static void appendText(XWPFRun run, String text) {
        int pos = run.getCTR().sizeOfTArray();
        run.setText(text, pos);
    }

    /** Убирает весь текст из run, оставляя только его форматирование (для превращения " " в разрыв строки). */
    private static void clearRunText(XWPFRun run) {
        CTR ctr = run.getCTR();
        while (ctr.sizeOfTArray() > 0) {
            ctr.removeT(0);
        }
    }

    // ---- Автоподбор размера шрифта под одну страницу ------------------------

    /**
     * После сборки документа уменьшает шрифт в столбце "Чтим память", пока
     * расписание не поместится на одну страницу (проверяется через реальную
     * конвертацию в PDF). Жирная "шапка" дня (см. {@link MemoryCell#protectedRuns})
     * в уменьшении не участвует.
     */
    private static final class PageFitter {

        private PageFitter() {
        }

        static void fit(XWPFDocument document, Path outputPath, List<MemoryCell> memoryCells)
                throws IOException {
            List<Integer> order = new ArrayList<>();
            for (int i = 0; i < memoryCells.size(); i++) {
                order.add(i);
            }
            // начинаем с самой длинной ячейки, дальше по убыванию длины текста
            order.sort((a, b) -> Integer.compare(
                    memoryCells.get(b).approximateLength, memoryCells.get(a).approximateLength));

            int pointer = 0;
            for (int step = 0; step < MAX_SHRINK_STEPS; step++) {
                saveDocument(document, outputPath);

                int pages;
                try {
                    pages = countPdfPages(outputPath);
                } catch (IOException | InterruptedException e) {
                    System.out.println("Автоподбор шрифта пропущен (нужны soffice и pdfinfo в PATH): "
                            + e.getMessage());
                    return;
                }

                if (pages <= 1) {
                    System.out.println("Расписание уместилось на 1 страницу.");
                    return;
                }
                if (order.isEmpty()) {
                    System.out.println("Достигнут минимальный размер шрифта (" + MIN_FONT_SIZE_PT
                            + " pt), но страниц всё ещё " + pages + ".");
                    return;
                }

                int cellIndex = order.get(pointer % order.size());
                boolean reduced = decreaseFontSize(memoryCells.get(cellIndex), 1);
                if (!reduced) {
                    // эта ячейка уже на минимальном размере — исключаем её из круга
                    order.remove(Integer.valueOf(cellIndex));
                    continue;
                }
                pointer++;
            }
            System.out.println("Достигнут предел попыток автоподбора шрифта (" + MAX_SHRINK_STEPS + ").");
        }

        /** Уменьшает на stepPt все НЕ защищённые runs ячейки. Возвращает false, если уменьшать уже нечего. */
        private static boolean decreaseFontSize(MemoryCell memoryCell, int stepPt) {
            boolean reducedAny = false;
            for (XWPFParagraph paragraph : memoryCell.cell.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    if (memoryCell.protectedRuns.contains(run)) {
                        continue;
                    }
                    int current = run.getFontSize();
                    if (current <= 0) {
                        current = BODY_FONT_SIZE_PT;
                    }
                    if (current - stepPt < MIN_FONT_SIZE_PT) {
                        continue;
                    }
                    run.setFontSize(current - stepPt);
                    reducedAny = true;
                }
            }
            return reducedAny;
        }

        private static int countPdfPages(Path docxPath) throws IOException, InterruptedException {
            Path outDir = docxPath.toAbsolutePath().getParent();
            String baseName = stripExtension(docxPath.getFileName().toString());
            Path pdfPath = outDir.resolve(baseName + ".pdf");
            Files.deleteIfExists(pdfPath);

            runProcess("soffice", "--headless", "--convert-to", "pdf",
                    "--outdir", outDir.toString(), docxPath.toAbsolutePath().toString());

            if (!Files.exists(pdfPath)) {
                throw new IOException("LibreOffice не создал " + pdfPath);
            }

            String info = runProcess("pdfinfo", pdfPath.toAbsolutePath().toString());
            Matcher matcher = Pattern.compile("Pages:\\s*(\\d+)").matcher(info);
            if (matcher.find()) {
                return Integer.parseInt(matcher.group(1));
            }
            throw new IOException("Не удалось разобрать вывод pdfinfo: " + info);
        }

        private static String stripExtension(String fileName) {
            int dot = fileName.lastIndexOf('.');
            return dot < 0 ? fileName : fileName.substring(0, dot);
        }

        private static String runProcess(String... command) throws IOException, InterruptedException {
            ProcessBuilder builder = new ProcessBuilder(command);
            builder.redirectErrorStream(true);
            Process process = builder.start();

            StringBuilder output = new StringBuilder();
            try (BufferedReader reader = new BufferedReader(
                    new InputStreamReader(process.getInputStream(), StandardCharsets.UTF_8))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    output.append(line).append('\n');
                }
            }
            int exitCode = process.waitFor();
            if (exitCode != 0) {
                throw new IOException(command[0] + " завершился с кодом " + exitCode + ": " + output);
            }
            return output.toString();
        }
    }
}