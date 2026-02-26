package Postclient;

import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;

import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import javax.swing.TransferHandler;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.stream.Collectors;

import jakarta.mail.*;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeBodyPart;
import jakarta.mail.internet.MimeMessage;
import jakarta.mail.internet.MimeMultipart;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.ss.usermodel.*;
import org.jetbrains.annotations.NotNull;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.DefaultCaret;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.*;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

public class PostClient {

    // ---------- ТЕКСТЫ ----------
    private static String SUBJECT_DEFAULT = "Леко Стайл";

    private static String BODY_HEADER =
            "{name}, добрый день!\n\n" +
                    "В октябре мы можем предложить лучшие цены на продукты, которые вы ранее закупали или тестировали.\n\n" +
                    "В их числе:\n";

    private static String BODY_FOOTER =
            "\nЕсли какие-то из них актуальны — напишите, пожалуйста, подготовим индивидуальное предложение.\n\n" +
                    "Спасибо.\n{sender}\n{company}\n{email}\n{phone}\n";

    // ---------- НАСТРОЙКИ ПО УМОЛЧАНИЮ ----------
    private static String SENDER;
    private static String COMPANY;
    private static String FROM_EMAIL;
    private static String PHONE;
    private static String SMTP_SERVER;
    private static int SMTP_PORT = 587;
    private static String SMTP_LOGIN;
    private static String SMTP_PASSWORD;
    private static String COPY_EMAIL;

    private static List<Draft> loadDraftsCsv(File csvFile) throws IOException {
        if (csvFile == null || !csvFile.exists()) {
            throw new FileNotFoundException("drafts.csv не найден: " + (csvFile == null ? "null" : csvFile.getAbsolutePath()));
        }

        // читаем bytes и убираем BOM если есть
        byte[] bytes = Files.readAllBytes(csvFile.toPath());
        int offset = 0;
        if (bytes.length >= 3
                && (bytes[0] & 0xFF) == 0xEF
                && (bytes[1] & 0xFF) == 0xBB
                && (bytes[2] & 0xFF) == 0xBF) {
            offset = 3;
        }

        try (Reader r = new InputStreamReader(new ByteArrayInputStream(bytes, offset, bytes.length - offset), StandardCharsets.UTF_8);
             CSVParser parser = CSVFormat.DEFAULT
                     .builder()
                     .setHeader()
                     .setSkipHeaderRecord(true)
                     .build()
                     .parse(r)) {

            // Ожидаемые колонки: To, Subject, Body
            List<Draft> out = new ArrayList<>();
            for (CSVRecord rec : parser) {
                String to = rec.isMapped("To") ? rec.get("To").trim() : "";
                String subject = rec.isMapped("Subject") ? rec.get("Subject").trim() : "";
                String body = rec.isMapped("Body") ? rec.get("Body") : "";

                if (to.isBlank()) continue; // пропускаем пустые строки
                if (subject.isBlank()) subject = SUBJECT_DEFAULT;

                out.add(new Draft(to, subject, body));
            }
            return out;
        }
    }

    // Drag-and-drop
    private static void enableFileDrop(JTextField field, boolean allowMultiple) {
        field.setTransferHandler(new TransferHandler() {
            @Override
            public boolean canImport(TransferSupport support) {
                if (!support.isDrop()) return false;
                return support.isDataFlavorSupported(DataFlavor.javaFileListFlavor);
            }

            @Override
            @SuppressWarnings("unchecked")
            public boolean importData(TransferSupport support) {
                if (!canImport(support)) return false;
                try {
                    Transferable t = support.getTransferable();
                    List<File> files =
                            (List<File>) t.getTransferData(DataFlavor.javaFileListFlavor);
                    if (files == null || files.isEmpty()) return false;

                    if (!allowMultiple) {
                        // Берём только первый файл
                        field.setText(files.get(0).getAbsolutePath());
                    } else {
                        // Несколько файлов – склеиваем в одну строку через ';'
                        String text = files.stream()
                                .map(File::getAbsolutePath)
                                .collect(Collectors.joining(";"));
                        field.setText(text);
                    }
                    return true;
                } catch (Exception ex) {
                    ex.printStackTrace();
                    return false;
                }
            }
        });

        // Для удобства: подсказка и возможность тащить текст из поля
        field.setToolTipText("Можно перетащить файл(ы) мышкой");
        field.setDragEnabled(true);
    }

    // ---------- МОДЕЛИ ----------
    static class RowMap {
        final Map<String, Object> map = new LinkedHashMap<>();

        Object get(String key) {
            return map.get(key);
        }

        String getStr(String key) {
            Object v = map.get(key);
            return v == null ? "" : v.toString();
        }

        void put(String key, Object val) {
            map.put(key, val);
        }
    }

    static class Draft {
        String to;
        String subject;
        String body;

        Draft(String to, String subject, String body) {
            this.to = to;
            this.subject = subject;
            this.body = body;
        }
    }

    // ---------- УТИЛИТЫ ----------
    private static String getenv(String key, String def) {
        String v = System.getenv(key);
        return v != null && !v.isEmpty() ? v : def;
    }

    private static boolean isExcel(String path) {
        String p = path.toLowerCase(Locale.ROOT);
        return p.endsWith(".xlsx") || p.endsWith(".xls");
    }

    private static Workbook openWorkbook(File f) throws IOException {
        try (InputStream in = new FileInputStream(f)) {
            return WorkbookFactory.create(in); // автоматически определяет .xls и .xlsx
        } catch (Exception e) {
            throw new IOException("Не удалось открыть Excel '" + f.getName() + "': " + e.getMessage(), e);
        }
    }

    private static String autodetectSheet(File f) throws IOException {
        try (Workbook wb = openWorkbook(f)) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                Sheet s = wb.getSheetAt(i);
                Row header = s.getRow(s.getFirstRowNum());
                if (header != null) {
                    int cols = header.getLastCellNum();
                    if (cols >= 3) return s.getSheetName();
                }
            }
            return wb.getSheetAt(0).getSheetName();
        } catch (Exception e) {
            throw new IOException("Не удалось открыть Excel '" + f.getName() + "': " + e.getMessage(), e);
        }
    }

    private static Map<String, String> normalizeColumns(List<String> cols) {
        Map<String, String> mapping = new LinkedHashMap<>();
        for (String c : cols) {
            String cl = (c == null ? "" : c).trim().toLowerCase(Locale.ROOT);
            String key;
            if (cl.contains("клиент") || cl.contains("client")) {
                key = "Клиент";
            } else if (cl.contains("артикул") || cl.contains("sku") || cl.contains("код")) {
                key = "Артикул";
            } else if (cl.contains("номенклат") || cl.contains("product") || cl.contains("наименование")
                    || cl.contains("товар") || cl.contains("позиция")) {
                key = "Номенклатура";
            } else if (cl.contains("стоим") || cl.contains("сумм") || cl.contains("amount")
                    || cl.contains("price") || cl.contains("руб")) {
                key = "Стоимость продажи (руб.)";
            } else if (cl.contains("дата") || cl.contains("date")) {
                key = "Дата";
            } else {
                key = c;
            }
            mapping.put(c, key);
        }
        return mapping;
    }

    private static List<RowMap> readSheet(File f, String sheetName) throws IOException {
        try (Workbook wb = openWorkbook(f)) {
            Sheet s = wb.getSheet(sheetName);
            if (s == null) throw new IOException("Лист не найден: " + sheetName);

            int first = s.getFirstRowNum();
            int last = s.getLastRowNum();
            if (first > last) return Collections.emptyList();

            // заголовки
            Row hdr = s.getRow(first);
            if (hdr == null) return Collections.emptyList();
            List<String> rawHeaders = new ArrayList<>();
            int maxCol = hdr.getLastCellNum();
            for (int c = 0; c < maxCol; c++) {
                rawHeaders.add(getCellString(hdr.getCell(c)));
            }
            Map<String, String> map = normalizeColumns(rawHeaders);

            // данные
            List<RowMap> rows = new ArrayList<>();
            for (int r = first + 1; r <= last; r++) {
                Row row = s.getRow(r);
                if (row == null) continue;
                RowMap rm = new RowMap();
                int cells = Math.max(maxCol, row.getLastCellNum());
                for (int c = 0; c < cells; c++) {
                    String srcKey = c < rawHeaders.size() ? rawHeaders.get(c) : ("COL" + c);
                    String normKey = map.getOrDefault(srcKey, srcKey);
                    rm.put(normKey, getCellValue(row.getCell(c)));
                }
                rows.add(rm);
            }
            return rows;
        }
    }

    private static String getCellString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    yield new SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue());
                } else {
                    double d = cell.getNumericCellValue();
                    if (Math.floor(d) == d) yield String.valueOf((long) d);
                    yield String.valueOf(d);
                }
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }

    private static Object getCellValue(Cell cell) {
        if (cell == null) return null;
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> (DateUtil.isCellDateFormatted(cell) ? cell.getDateCellValue() : cell.getNumericCellValue());
            default -> null;
        };
    }

    private static List<RowMap> loadPivot(File pivot) throws IOException {
        String sheet = autodetectSheet(pivot);
        List<RowMap> rows = readSheet(pivot, sheet);

        // Проверки обязательных колонок
        Set<String> cols = collectColumns(rows);
        for (String col : List.of("Клиент", "Артикул", "Номенклатура")) {
            if (!cols.contains(col)) throw new IOException("В Excel не найдена обязательная колонка: " + col);
        }
        if (!cols.contains("Стоимость продажи (руб.)")) {
            for (RowMap r : rows) r.put("Стоимость продажи (руб.)", null);
        }
        if (!cols.contains("Дата")) {
            for (RowMap r : rows) r.put("Дата", null);
        }

        // Трим строк
        for (RowMap r : rows) {
            for (String k : List.of("Клиент", "Артикул", "Номенклатура")) {
                r.put(k, r.getStr(k).trim());
            }
        }
        return rows;
    }

    private static Set<String> collectColumns(List<RowMap> rows) {
        Set<String> keys = new LinkedHashSet<>();
        for (RowMap r : rows) keys.addAll(r.map.keySet());
        return keys;
    }

    private static List<RowMap> loadContacts(File contacts) throws IOException {
        String sheet = autodetectSheet(contacts);
        List<RowMap> rows = readSheet(contacts, sheet);
        Set<String> cols = collectColumns(rows);
        for (String col : List.of("Клиент", "Имя", "Email")) {
            if (!cols.contains(col)) throw new IOException("В '" + contacts.getName() + "' не найдена колонка: " + col);
        }
        for (RowMap r : rows) {
            for (String k : List.of("Клиент", "Имя", "Email")) {
                r.put(k, r.getStr(k).trim());
            }
        }
        return rows;
    }

    private static String replacePlaceholders(String template, Map<String, String> vals) {
        String out = template;
        for (Map.Entry<String, String> e : vals.entrySet()) {
            out = out.replace("{" + e.getKey() + "}", e.getValue() == null ? "" : e.getValue());
        }
        return out;
    }

    private static String buildBody(String name, List<String[]> items,
                                    String sender, String company, String email, String phone) {
        StringBuilder lines = new StringBuilder();
        for (String[] it : items) {
            String art = (it[0] == null ? "" : it[0].trim());
            String title = (it[1] == null ? "" : it[1].trim());
            if (art.isEmpty() && title.isEmpty()) continue;
            if (!art.isEmpty() && !title.isEmpty()) {
                lines.append("- ").append(art).append(" — ").append(title).append("\n");
            } else if (!title.isEmpty()) {
                lines.append("- ").append(title).append("\n");
            } else {
                lines.append("- ").append(art).append("\n");
            }
        }

        Map<String, String> ph = new HashMap<>();
        ph.put("name", name);
        ph.put("sender", sender);
        ph.put("company", company);
        ph.put("email", email);
        ph.put("phone", phone);

        return replacePlaceholders(BODY_HEADER, ph) + lines + replacePlaceholders(BODY_FOOTER, ph);
    }

    private static Date tryParseDate(Object v) {
        if (v == null) return null;
        if (v instanceof Date) return (Date) v;
        String s = v.toString().trim();
        if (s.isEmpty()) return null;
        // Несколько форматов
        String[] fmts = {"yyyy-MM-dd", "dd.MM.yyyy", "dd/MM/yyyy", "yyyy/MM/dd", "dd-MM-yyyy"};
        for (String f : fmts) {
            try {
                return new SimpleDateFormat(f).parse(s);
            } catch (ParseException ignored) {
            }
        }
        return null;
    }

    private static List<Draft> makeDrafts(List<RowMap> pivot, List<RowMap> contacts, int maxItems) {
        List<Draft> out = new ArrayList<>();

        // Индекс по Клиенту
        Map<String, List<RowMap>> byClient = new HashMap<>();
        for (RowMap r : pivot) {
            String client = r.getStr("Клиент");
            byClient.computeIfAbsent(client.toUpperCase(Locale.ROOT), k -> new ArrayList<>()).add(r);
        }

        for (RowMap c : contacts) {
            String client = c.getStr("Клиент");
            String name = c.getStr("Имя");
            String email = c.getStr("Email");

            List<RowMap> subset = byClient.getOrDefault(client.toUpperCase(Locale.ROOT), new ArrayList<>());
            // unique по (Артикул, Номенклатура)
            LinkedHashSet<String> seen = new LinkedHashSet<>();
            List<RowMap> dedup = new ArrayList<>();
            for (RowMap r : subset) {
                String key = r.getStr("Артикул") + "||" + r.getStr("Номенклатура");
                if (seen.add(key)) dedup.add(r);
            }
            // сортировка по "Дата" (desc)
            dedup.sort((a, b) -> {
                Date da = tryParseDate(a.get("Дата"));
                Date db = tryParseDate(b.get("Дата"));
                if (da == null && db == null) return 0;
                if (da == null) return 1;
                if (db == null) return -1;
                return -da.compareTo(db);
            });

            List<String[]> items = new ArrayList<>();
            for (RowMap r : dedup) {
                items.add(new String[]{r.getStr("Артикул"), r.getStr("Номенклатура")});
                if (items.size() >= maxItems) break;
            }

            String body;
            if (items.isEmpty()) {
                body = buildBody(name,
                        Collections.singletonList(
                                new String[]{"", "— данных по продуктам нет —"}
                        ),
                        SENDER, COMPANY, FROM_EMAIL, PHONE);
            } else {
                body = buildBody(name, items, SENDER, COMPANY, FROM_EMAIL, PHONE);
            }
            out.add(new Draft(email, SUBJECT_DEFAULT, body));
        }
        return out;
    }

    private static void saveDraftsCsv(List<Draft> drafts, File outFile) throws IOException {
        // UTF-8 BOM + CSV
        try (OutputStream os = new FileOutputStream(outFile)) {
            os.write(new byte[]{(byte) 0xEF, (byte) 0xBB, (byte) 0xBF});
        }
        try (Writer w = Files.newBufferedWriter(outFile.toPath(), StandardCharsets.UTF_8, StandardOpenOption.APPEND);
             CSVPrinter csv = new CSVPrinter(w, CSVFormat.DEFAULT.builder().setHeader("To", "Subject", "Body").build())) {
            for (Draft d : drafts) {
                csv.printRecord(d.to, d.subject, d.body);
            }
        }
    }

    private static void sendEmails(List<Draft> drafts,
                                   String server, int port,
                                   String login, String password,
                                   String fromEmail,
                                   List<File> attachmentFiles,
                                   boolean hasPivot,
                                   ProgressCb progress,
                                   LogCb logger) throws MessagingException {

        Properties props = new Properties();
        props.put("mail.smtp.host", SMTP_SERVER);
        props.put("mail.smtp.port", SMTP_PORT);
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.starttls.required", "true");
        props.put("mail.smtp.ssl.protocols", "TLSv1.2");
        props.put("mail.smtp.ssl.trust", SMTP_SERVER);
        props.put("mail.smtp.ssl.checkserveridentity", "true");

        Session session = Session.getInstance(props, new Authenticator() {
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(SMTP_LOGIN, SMTP_PASSWORD);
            }
        });

        session.setDebug(true);

        Transport transport = session.getTransport("smtp");
        int total = drafts.size();

        int index = -1;

        try {

            transport.connect(); // EHLO -> STARTTLS -> TLS
            int sent = 0;

            Path pathStaticFile = getFileProgramData();

            try (BufferedWriter writer = Files.newBufferedWriter(
                    pathStaticFile,
                    StandardOpenOption.CREATE,
                    StandardOpenOption.APPEND))
            {

                // Заполнение начальных данных перед отправкой
                writer.newLine();
                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
                String formattedTime = LocalDateTime.now().format(formatter);
                writer.write("Начало рассылки:" + formattedTime);
                writer.newLine();
                writer.write(System.getProperty("user.name") + props.toString());
                writer.newLine();

                for (Draft d : drafts) {

                    // ВАЖНО: теперь "Пропущено (нет данных)" только если pivot реально был указан,
                    if ( hasPivot && d.body.contains("— данных по продуктам нет —")) {
                        if (logger != null) {
                            logger.log("Пропущено (нет данных): " + d.to);
                        }
                        continue;
                    } else {
                        d.body = d.body.replace("— данных по продуктам нет —", "   ");
                    }

                    MimeMessage msg = new MimeMessage(session);
                    msg.setFrom(new InternetAddress(login));
                    msg.setRecipients(Message.RecipientType.TO, InternetAddress.parse(d.to, false));

                    if (COPY_EMAIL != null && !COPY_EMAIL.isBlank()) {
                        msg.addRecipients(Message.RecipientType.BCC, InternetAddress.parse(COPY_EMAIL, false));
                    }

                    msg.setSubject(d.subject, "UTF-8");

                    // --- Тут прикрепляем файл(ы), если они есть ---
                    if (attachmentFiles != null && !attachmentFiles.isEmpty()) {
                        MimeBodyPart textPart = new MimeBodyPart();
                        textPart.setText(d.body, "UTF-8");

                        Multipart multipart = new MimeMultipart();
                        multipart.addBodyPart(textPart);

                        for (File att : attachmentFiles) {
                            if (att == null) continue;
                            MimeBodyPart attachmentPart = new MimeBodyPart();
                            try {
                                attachmentPart.attachFile(att);
                                multipart.addBodyPart(attachmentPart);
                            } catch (IOException e) {
                                if (logger != null) {
                                    logger.log("[Ошибка] Не удалось прикрепить файл '" +
                                            att.getAbsolutePath() + "': " + e.getMessage());
                                }
                            }
                        }

                        msg.setContent(multipart);
                    } else {
                        // без вложений — как раньше
                        msg.setText(d.body, "UTF-8");
                    }

                    transport.sendMessage(msg, msg.getAllRecipients());

                    sent++;
                    if (logger != null) logger.log("Отправлено: " + d.to + " (" + sent + "/" + total + ")");
                    if (progress != null) progress.onProgress(sent, total);

//                    String line =
                    writer.write("Отправлено: " + d.to + " (" + sent + "/" + total + ")");
                    writer.newLine();

                    try {
                        Thread.sleep(50);
                    } catch (InterruptedException ignored) {
                    }

                    if(sent % 25 == 0) {
                        transport.close();
                        transport.connect();
                    }

                    index = drafts.indexOf(d);

                }
            }  catch (IOException e) {
                e.printStackTrace();
            }
        } catch (AuthenticationFailedException afe) {

            throw new RuntimeException("Аутентификация отклонена сервером. " +
                    "Проверьте логин/пароль и совпадение From с логином. См. SMTP debug выше.", afe);

        } finally {

            if(index >= 0) {
                drafts.subList(0, index).clear();
            }

            if (drafts.size() > 0) {
                // Сохраним черновик
                File out = new File(getPathToFolderSetting() + "/drafts.csv");
                try {
                    saveDraftsCsv(drafts, out);
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
                SwingUtilities.invokeLater(() -> {
                    logger.log("Черновики сохранены.");
                });
            }

            try {
                transport.close();
            } catch (Exception ignored) {
            }
        }
    }

    interface ProgressCb {
        void onProgress(int done, int total);
    }

    interface LogCb {
        void log(String msg);
    }

    // ---------- CLI ----------
    private static void runCli(Map<String, String> args) {
        try {
            File pivot = new File(reqArg(args, "--pivot"));
            File contacts = new File(reqArg(args, "--contacts"));
            int maxItems = Integer.parseInt(args.getOrDefault("--max_items", "50"));
            boolean send = args.containsKey("--send");

            SUBJECT_DEFAULT = args.getOrDefault("--subject", SUBJECT_DEFAULT);
            SMTP_SERVER = args.getOrDefault("--smtp_server", SMTP_SERVER);
            SMTP_PORT = Integer.parseInt(args.getOrDefault("--smtp_port", String.valueOf(SMTP_PORT)));
            SMTP_LOGIN = args.getOrDefault("--smtp_login", SMTP_LOGIN);
            FROM_EMAIL = args.getOrDefault("--from_email", FROM_EMAIL);

            List<RowMap> pivotDf = loadPivot(pivot);
            List<RowMap> contactsDf = loadContacts(contacts);

            List<Draft> drafts = makeDrafts(pivotDf, contactsDf, maxItems);

            if (send) {
                String pwd = System.getenv().getOrDefault("SMTP_PASSWORD", SMTP_PASSWORD);
                if (pwd == null || pwd.isEmpty()) {
                    System.err.println("[!] Не задан SMTP_PASSWORD. Пример:\n" +
                            "SMTP_PASSWORD=*** java -jar bulk-mailer.jar --contacts ... --pivot ... --send");
                    System.exit(1);
                }

                sendEmails(
                        drafts,
                        SMTP_SERVER, SMTP_PORT,
                        SMTP_LOGIN, pwd,
                        FROM_EMAIL,
                        Collections.emptyList(),              // вложения для CLI нет
                        true,              // в CLI pivot всегда обязателен
                        (done, total) -> {
                        },
                        System.out::println
                );

                System.out.println("Готово: отправлено " + drafts.size() + " писем.");
            } else {
                File out = new File(getPathToFolderSetting() + "/drafts.csv");
                saveDraftsCsv(drafts, out);
                System.out.println("Создан файл drafts.csv (черновики писем).");
            }
        } catch (Exception e) {
            System.err.println("[!] Ошибка: " + e.getMessage());
            System.exit(1);
        }
    }

    private static String reqArg(Map<String, String> args, String key) {
        String v = args.get(key);
        if (v == null || v.isEmpty()) throw new IllegalArgumentException("Отсутствует аргумент " + key);
        return v;
    }

    private static Map<String, String> parseArgs(String[] argv) {
        Map<String, String> map = new LinkedHashMap<>();
        for (int i = 0; i < argv.length; i++) {
            String a = argv[i];
            if (a.startsWith("--")) {
                if ((i + 1) < argv.length && !argv[i + 1].startsWith("--")) {
                    map.put(a, argv[++i]);
                } else {
                    map.put(a, "");
                }
            }
        }
        return map;
    }

    // ---------- GUI (Swing) ----------
    private static void runGui() {
        SwingUtilities.invokeLater(() -> {
            JFrame f = new JFrame("Leko Style Mailer — рассылка");
            f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            f.setMinimumSize(new Dimension(1000, 800));
            f.setSize(1100, 900);

            // Добавил событие закрытия формы для сохранения настроек (логины, пароли...)
            f.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);

            f.addWindowListener(new WindowAdapter() {
                @Override
                public void windowClosing(WindowEvent e) {
                    saveSettings(f);
                    f.dispose();   // закрываем окно
                    System.exit(0); // если это главное окно
                }
            });

            JPanel root = new JPanel();
            root.setLayout(new BorderLayout(10, 10));
            root.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
            f.setContentPane(root);

            // Файлы
            JPanel files = new JPanel(new GridBagLayout());
            files.setBorder(titled("Файлы"));

            JTextField tfContacts = new JTextField();
            JButton btnContacts = new JButton("Выбрать…");
            JTextField tfPivot = new JTextField();
            JButton btnPivot = new JButton("Выбрать…");
            JTextField tfAttachment = new JTextField();
            JButton btnAttachment = new JButton("Выбрать…");

            JTextField tfDrafts = new JTextField();
            JButton btnDrafts = new JButton("Выбрать…");
            enableFileDrop(tfDrafts, false);

            // Drag-and-drop для выбора файлов
            enableFileDrop(tfContacts, false);    // один contacts файл
            enableFileDrop(tfPivot, false);       // один pivot файл
            enableFileDrop(tfAttachment, true);   // можно перетащить несколько вложений

            GridBagConstraints c = new GridBagConstraints();
            c.insets = new Insets(6, 8, 6, 8);
            c.gridx = 0;
            c.gridy = 0;
            c.anchor = GridBagConstraints.WEST;
            files.add(new JLabel("Contacts файл:"), c);
            c.gridx = 1;
            c.weightx = 1;
            c.fill = GridBagConstraints.HORIZONTAL;
            files.add(tfContacts, c);
            c.gridx = 2;
            c.weightx = 0;
            c.fill = GridBagConstraints.NONE;
            files.add(btnContacts, c);

            c.gridx = 0;
            c.gridy = 1;
            c.anchor = GridBagConstraints.WEST;
            files.add(new JLabel("Pivot файл:"), c);
            c.gridx = 1;
            c.weightx = 1;
            c.fill = GridBagConstraints.HORIZONTAL;
            files.add(tfPivot, c);
            c.gridx = 2;
            c.weightx = 0;
            c.fill = GridBagConstraints.NONE;
            files.add(btnPivot, c);

            c.gridx = 0;
            c.gridy = 2;
            c.anchor = GridBagConstraints.WEST;
            files.add(new JLabel("Drafts CSV (опционально):"), c);
            c.gridx = 1;
            c.weightx = 1;
            c.fill = GridBagConstraints.HORIZONTAL;
            files.add(tfDrafts, c);
            c.gridx = 2;
            c.weightx = 0;
            c.fill = GridBagConstraints.NONE;
            files.add(btnDrafts, c);

            c.gridx = 0;
            c.gridy = 3;
            c.anchor = GridBagConstraints.WEST;
            files.add(new JLabel("Вложение (опционально):"), c);
            c.gridx = 1;
            c.weightx = 1;
            c.fill = GridBagConstraints.HORIZONTAL;
            files.add(tfAttachment, c);
            c.gridx = 2;
            c.weightx = 0;
            c.fill = GridBagConstraints.NONE;
            files.add(btnAttachment, c);

            // --- Панель "Файлы" уже создана выше ---
            // Теперь создаём панель с логотипом справа от "Файлы"
            JLabel logoLabel = null;
            try {
                // Если используешь resources-папку:
                URL logoUrl = PostClient.class.getResource("/com/mycompany/postclient/logo.png");
                if (logoUrl != null) {
                    ImageIcon logoIcon = new ImageIcon(logoUrl);
                    Image scaled = logoIcon.getImage().getScaledInstance(56, 55, Image.SCALE_SMOOTH);
                    logoLabel = new JLabel(new ImageIcon(scaled));
                } else {
                    System.err.println("Не найден /com/mycompany/postclient/logo.png");
                }
            } catch (Exception ex) {
                System.err.println("Ошибка при загрузке логотипа: " + ex.getMessage());
            }

            // Оборачиваем панель "Файлы" и логотип в одну горизонтальную панель
            JPanel filesRow = new JPanel(new BorderLayout(10, 0));
            filesRow.add(files, BorderLayout.CENTER);

            if (logoLabel != null) {
                JPanel logoPanel = new JPanel(new BorderLayout());
                logoPanel.setOpaque(false);
                logoPanel.add(logoLabel, BorderLayout.NORTH); // прижать к верху
                filesRow.add(logoPanel, BorderLayout.EAST);
            }

            // Тексты и логи
            JPanel texts = new JPanel(new GridBagLayout());
            texts.setBorder(titled("Тексты письма и логи"));

            JLabel lbPh = new JLabel("Плейсхолдеры: {name}, {sender}, {company}, {email}, {phone}");

            JTextArea taHeader = new JTextArea(4, 80);
            taHeader.setLineWrap(true);
            taHeader.setWrapStyleWord(true);
            taHeader.setText(BODY_HEADER);

            JTextArea taFooter = new JTextArea(4, 80);
            taFooter.setLineWrap(true);
            taFooter.setWrapStyleWord(true);
            taFooter.setText(BODY_FOOTER);

            JTextArea taLogs = new JTextArea(10, 80);
            taLogs.setEditable(false);
            DefaultCaret caret = (DefaultCaret) taLogs.getCaret();
            caret.setUpdatePolicy(DefaultCaret.ALWAYS_UPDATE);
            JScrollPane spHeader = new JScrollPane(taHeader);
            JScrollPane spFooter = new JScrollPane(taFooter);
            JScrollPane spLogs = new JScrollPane(taLogs);

            int gy = 0;
            c = new GridBagConstraints();
            c.insets = new Insets(6, 8, 6, 8);
            c.gridx = 0;
            c.gridy = gy;
            c.gridwidth = 2;
            c.anchor = GridBagConstraints.WEST;
            texts.add(lbPh, c);
            gy++;

            c.gridx = 0;
            c.gridy = gy;
            c.gridwidth = 1;
            c.anchor = GridBagConstraints.NORTHWEST;
            texts.add(new JLabel("BODY_HEADER:"), c);
            c.gridx = 1;
            c.gridy = gy;
            c.weightx = 1;
            c.weighty = 1;
            c.fill = GridBagConstraints.BOTH;
            texts.add(spHeader, c);
            gy++;

            c.gridx = 0;
            c.gridy = gy;
            c.weightx = 0;
            c.weighty = 0;
            c.fill = GridBagConstraints.NONE;
            c.anchor = GridBagConstraints.NORTHWEST;
            texts.add(new JLabel("BODY_FOOTER:"), c);
            c.gridx = 1;
            c.gridy = gy;
            c.weightx = 1;
            c.weighty = 1;
            c.fill = GridBagConstraints.BOTH;
            texts.add(spFooter, c);
            gy++;

            c.gridx = 0;
            c.gridy = gy;
            c.weightx = 0;
            c.weighty = 0;
            c.fill = GridBagConstraints.NONE;
            c.anchor = GridBagConstraints.NORTHWEST;
            texts.add(new JLabel("Логи:"), c);
            c.gridx = 1;
            c.gridy = gy;
            c.weightx = 1;
            c.weighty = 2;
            c.fill = GridBagConstraints.BOTH;
            texts.add(spLogs, c);

            // Прочитаем файл с настройками если они есть
            loadSettings();

            // Параметры
            JPanel params = new JPanel(new GridBagLayout());
            params.setBorder(titled("Параметры писем и SMTP"));

            JTextField tfSubject = new JTextField(SUBJECT_DEFAULT);
            JSpinner spMaxItems = new JSpinner(new SpinnerNumberModel(50, 1, 100, 1));
            JTextField tfSender = new JTextField(SENDER);
            JTextField tfCompany = new JTextField(COMPANY);
            JTextField tfFrom = new JTextField(FROM_EMAIL);
            JTextField tfPhone = new JTextField(PHONE);
            JTextField tfSmtpServer = new JTextField(SMTP_SERVER);
            JTextField tfCopy = new JTextField(COPY_EMAIL);
            JSpinner spSmtpPort = new JSpinner(new SpinnerNumberModel(SMTP_PORT, 1, 65535, 1));
            JTextField tfSmtpLogin = new JTextField(SMTP_LOGIN);
            JPasswordField pfSmtpPass = new JPasswordField(SMTP_PASSWORD);

            JCheckBox cbSend = new JCheckBox("Сразу отправлять письма (иначе — сохранить drafts.csv)", true);

            int r = 0;
            c = new GridBagConstraints();
            c.insets = new Insets(6, 8, 6, 8);
            c.gridx = 0;
            c.gridy = r;
            c.anchor = GridBagConstraints.WEST;
            params.add(new JLabel("Тема:"), c);
            c.gridx = 1;
            c.gridy = r;
            c.weightx = 1;
            c.fill = GridBagConstraints.HORIZONTAL;
            params.add(tfSubject, c);
            c.gridx = 2;
            c.gridy = r;
            c.anchor = GridBagConstraints.EAST;
            params.add(new JLabel("Max позиций:"), c);
            c.gridx = 3;
            c.gridy = r;
            params.add(spMaxItems, c);
            r++;

            c.gridx = 0;
            c.gridy = r;
            c.anchor = GridBagConstraints.EAST;
            params.add(new JLabel("Отправитель (Имя):"), c);
            c.gridx = 1;
            c.gridy = r;
            c.weightx = 1;
            c.fill = GridBagConstraints.HORIZONTAL;
            params.add(tfSender, c);
            c.gridx = 2;
            c.gridy = r;
            c.anchor = GridBagConstraints.EAST;
            params.add(new JLabel("Компания:"), c);
            c.gridx = 3;
            c.gridy = r;
            c.fill = GridBagConstraints.HORIZONTAL;
            params.add(tfCompany, c);
            r++;

            c.gridx = 0;
            c.gridy = r;
            c.anchor = GridBagConstraints.EAST;
            params.add(new JLabel("From email:"), c);
            c.gridx = 1;
            c.gridy = r;
            c.fill = GridBagConstraints.HORIZONTAL;
            params.add(tfFrom, c);

            c.gridx = 2;
            c.gridy = r;
            c.anchor = GridBagConstraints.EAST;
            params.add(new JLabel("Телефон:"), c);
            c.gridx = 3;
            c.gridy = r;
            c.fill = GridBagConstraints.HORIZONTAL;
            params.add(tfPhone, c);
            r++;

            c.gridx = 0;
            c.gridy = r;
            c.anchor = GridBagConstraints.EAST;
            params.add(new JLabel("SMTP сервер:"), c);
            c.gridx = 1;
            c.gridy = r;
            c.fill = GridBagConstraints.HORIZONTAL;
            params.add(tfSmtpServer, c);
            c.gridx = 2;
            c.gridy = r;
            c.anchor = GridBagConstraints.EAST;
            params.add(new JLabel("Порт:"), c);
            c.gridx = 3;
            c.gridy = r;
            params.add(spSmtpPort, c);
            r++;

            c.gridx = 0;
            c.gridy = r;
            c.anchor = GridBagConstraints.EAST;
            params.add(new JLabel("SMTP логин:"), c);
            c.gridx = 1;
            c.gridy = r;
            c.fill = GridBagConstraints.HORIZONTAL;
            params.add(tfSmtpLogin, c);
            c.gridx = 2;
            c.gridy = r;
            c.anchor = GridBagConstraints.EAST;
            params.add(new JLabel("SMTP пароль:"), c);
            c.gridx = 3;
            c.gridy = r;
            c.fill = GridBagConstraints.HORIZONTAL;
            params.add(pfSmtpPass, c);
            r++;

            c.gridx = 0;
            c.gridy = r;
            c.anchor = GridBagConstraints.EAST;
            params.add(new JLabel("Copy email:"), c);
            c.gridx = 1;
            c.gridy = r;
            c.fill = GridBagConstraints.HORIZONTAL;
            params.add(tfCopy, c);
            r++;

            c.gridx = 0;
            c.gridy = r;
            c.gridwidth = 4;
            c.anchor = GridBagConstraints.WEST;
            params.add(cbSend, c);

            // Запуск
            JPanel run = new JPanel(new BorderLayout(8, 8));
            run.setBorder(titled("Запуск"));
            JProgressBar pb = new JProgressBar(0, 1);
            JPanel buttons = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
            JButton btnRun = new JButton("Запустить");
            JButton btnStop = new JButton("Стоп");
            btnStop.setEnabled(false);
            JLabel lbStatus = new JLabel("Готово");
            buttons.add(btnRun);
            buttons.add(btnStop);
            run.add(pb, BorderLayout.NORTH);
            run.add(buttons, BorderLayout.CENTER);
            run.add(lbStatus, BorderLayout.SOUTH);

            // Сборка
            JPanel center = new JPanel();
            center.setLayout(new BoxLayout(center, BoxLayout.Y_AXIS));
            center.add(filesRow); //files
            center.add(Box.createVerticalStrut(8));
            center.add(texts);
            center.add(Box.createVerticalStrut(8));
            center.add(params);
            center.add(Box.createVerticalStrut(8));
            center.add(run);
            root.add(center, BorderLayout.CENTER);

            // Файловые диалоги
            btnContacts.addActionListener(ev -> {
                JFileChooser ch = new JFileChooser();
                ch.setDialogTitle("Выберите contacts.xlsx");
                ch.setFileFilter(new FileNameExtensionFilter("Excel", "xlsx", "xls"));
                if (ch.showOpenDialog(f) == JFileChooser.APPROVE_OPTION) {
                    tfContacts.setText(ch.getSelectedFile().getAbsolutePath());
                }
            });
            btnPivot.addActionListener(ev -> {
                JFileChooser ch = new JFileChooser();
                ch.setDialogTitle("Выберите pivot.xlsx");
                ch.setFileFilter(new FileNameExtensionFilter("Excel", "xlsx", "xls"));
                if (ch.showOpenDialog(f) == JFileChooser.APPROVE_OPTION) {
                    tfPivot.setText(ch.getSelectedFile().getAbsolutePath());
                }
            });

            btnAttachment.addActionListener(ev -> {
                JFileChooser ch = new JFileChooser();
                ch.setDialogTitle("Выберите файл(ы) для вложения");
                ch.setMultiSelectionEnabled(true); // <-- ВАЖНО: разрешаем несколько

                if (ch.showOpenDialog(f) == JFileChooser.APPROVE_OPTION) {
                    File[] filesD = ch.getSelectedFiles();

                    if (filesD != null && filesD.length > 0) {
                        StringBuilder sb = new StringBuilder();
                        for (int i = 0; i < filesD.length; i++) {
                            if (i > 0) sb.append(";");
                            sb.append(filesD[i].getAbsolutePath());
                        }
                        tfAttachment.setText(sb.toString());
                    }
                }
            });

            btnDrafts.addActionListener(ev -> {
                JFileChooser ch = new JFileChooser();
                ch.setDialogTitle("Выберите drafts.csv");
                ch.setFileFilter(new FileNameExtensionFilter("CSV", "csv"));
                if (ch.showOpenDialog(f) == JFileChooser.APPROVE_OPTION) {
                    tfDrafts.setText(ch.getSelectedFile().getAbsolutePath());
                }
            });

            // Логгер
            LogCb logger = msg -> SwingUtilities.invokeLater(() -> {
                taLogs.append(msg + "\n");
            });


            // Кнопки
            btnRun.addActionListener((ActionEvent e) -> {

                String draftsPath = tfDrafts.getText().trim();      // <-- добавь
                String contactsPath = tfContacts.getText().trim();  // <-- можно сразу, чтобы ниже не дублировать

                // Если drafts.csv указан — contacts НЕ требуем
                if (draftsPath.isBlank() && contactsPath.isEmpty()) {
                    JOptionPane.showMessageDialog(f, "Укажите contacts.xlsx (или выберите drafts.csv).",
                            "Нужно выбрать файл", JOptionPane.WARNING_MESSAGE);
                    return;
                }

                // Применяем тексты и плейсхолдеры
                BODY_HEADER = taHeader.getText();
                BODY_FOOTER = taFooter.getText();
                try {
                    String test = replacePlaceholders(BODY_HEADER, Map.of(
                            "name", "Тест", "sender", tfSender.getText().trim(),
                            "company", tfCompany.getText().trim(), "email", tfFrom.getText().trim(),
                            "phone", tfPhone.getText().trim()));
                    test = replacePlaceholders(BODY_FOOTER, Map.of(
                            "name", "Тест", "sender", tfSender.getText().trim(),
                            "company", tfCompany.getText().trim(), "email", tfFrom.getText().trim(),
                            "phone", tfPhone.getText().trim()));
                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(f, "Ошибка плейсхолдеров: " + ex.getMessage(),
                            "Ошибка", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                SUBJECT_DEFAULT = tfSubject.getText().trim().isEmpty() ? SUBJECT_DEFAULT : tfSubject.getText().trim();
                SENDER = tfSender.getText().trim();
                COMPANY = tfCompany.getText().trim();
                FROM_EMAIL = tfFrom.getText().trim();
                PHONE = tfPhone.getText().trim();
                COPY_EMAIL = tfCopy.getText().trim();

                SMTP_SERVER = tfSmtpServer.getText().trim();
//                String smtpServer = SMTP_SERVER;
//                int smtpPort = (Integer) spSmtpPort.getValue();
                SMTP_PORT = (Integer) spSmtpPort.getValue();
                SMTP_LOGIN = tfSmtpLogin.getText().trim();
//                String smtpLogin = SMTP_LOGIN;
                SMTP_PASSWORD = new String(pfSmtpPass.getPassword());
//                String smtpPass = SMTP_PASSWORD;
                int maxItems = (Integer) spMaxItems.getValue();
                boolean doSend = cbSend.isSelected();

                String pivotPath = tfPivot.getText().trim();        // может быть пустым
                String attachmentPath = tfAttachment.getText().trim();   // может быть пустым

                taLogs.setText("");
                pb.setMaximum(1);
                pb.setValue(0);
                lbStatus.setText("Старт…");
                btnRun.setEnabled(false);
                btnStop.setEnabled(false);

                new Thread(() -> {
                    try {
                        SwingUtilities.invokeLater(() -> {
                            lbStatus.setText("Загрузка файлов…");
                            logger.log("Чтение contacts и (опционально) pivot…");
                        });

                        List<Draft> drafts;
                        boolean hasPivot;

                        if (!draftsPath.isBlank()) {
                            File df = new File(draftsPath);
                            drafts = loadDraftsCsv(df);
                            hasPivot = false; // чтобы НЕ срабатывала логика "пропущено (нет данных)"
                            logger.log("Загружен drafts.csv: " + df.getAbsolutePath());
                            logger.log("Писем в drafts.csv: " + drafts.size());

                        } else {
                            // contacts обязателен
                            List<RowMap> contactsDf = loadContacts(new File(contactsPath));

                            // pivot опционален
                            List<RowMap> pivotDf;
                            if (pivotPath.isBlank()) {
                                pivotDf = Collections.emptyList();
                                hasPivot = false;
                                logger.log("Pivot файл не указан — письма будут без списка продуктов.");
                            } else {
                                pivotDf = loadPivot(new File(pivotPath));
                                hasPivot = true;
                                logger.log("Загружен pivot: " + pivotPath);
                            }

                            SwingUtilities.invokeLater(() -> lbStatus.setText("Формирование черновиков…"));
                            drafts = makeDrafts(pivotDf, contactsDf, maxItems);
                            logger.log("Сформировано " + drafts.size() + " писем.");
                        }

                        // Подготовим файл вложения (если указан)
                        // Подготовим список файлов вложений (если указаны)
                        final List<File> attachmentFiles = new ArrayList<>();
                        if (attachmentPath != null && !attachmentPath.isBlank()) {
                            String[] parts = attachmentPath.split(";");
                            for (String p : parts) {
                                String trimmed = p.trim();
                                if (!trimmed.isEmpty()) {
                                    attachmentFiles.add(new File(trimmed));
                                }
                            }
                        }

                        if (doSend) {
                            if (SMTP_PASSWORD == null || SMTP_PASSWORD.isBlank()) {
                                throw new RuntimeException("SMTP пароль пуст. Укажите пароль или снимите галочку отправки.");
                            }
                            SwingUtilities.invokeLater(() -> {
                                lbStatus.setText("Отправка писем…");
                                pb.setMaximum(Math.max(drafts.size(), 1));
                            });
                            sendEmails(
                                    drafts,
                                    SMTP_SERVER, SMTP_PORT, SMTP_LOGIN, SMTP_PASSWORD, FROM_EMAIL,
                                    attachmentFiles, hasPivot,
                                    (done, total) -> SwingUtilities.invokeLater(() -> pb.setValue(done)),
                                    logger
                            );
                            SwingUtilities.invokeLater(() -> lbStatus.setText("Готово: отправлено."));
                        } else {
                            File out = new File(getPathToFolderSetting() + "/drafts.csv");
                            saveDraftsCsv(drafts, out);
                            SwingUtilities.invokeLater(() -> {
                                pb.setMaximum(1);
                                pb.setValue(1);
                                lbStatus.setText("Готово: сохранён drafts.csv");
                                logger.log("Черновики сохранены.");
                            });
                        }
                    } catch (Exception ex) {
                        logger.log("[Ошибка] " + ex.getMessage());
                        SwingUtilities.invokeLater(() -> lbStatus.setText("Ошибка"));
                    } finally {
                        SwingUtilities.invokeLater(() -> {
                            btnRun.setEnabled(true);
                            btnStop.setEnabled(false);
                        });
                    }

                }).start();
            });

            btnStop.addActionListener(ev -> JOptionPane.showMessageDialog(f,
                    "Остановка доступна после завершения текущего письма.", "Стоп", JOptionPane.INFORMATION_MESSAGE));
            f.setLocationRelativeTo(null);
            f.setVisible(true);
        });
    }

    private static void saveSettings(JFrame f) {
        Properties props = new Properties();

        String userName = System.getProperty("user.name");
        setIfNotNull(props, "USER", userName);
        setIfNotNull(props, "SENDER", SENDER);
        setIfNotNull(props, "COMPANY", COMPANY);
        setIfNotNull(props, "FROM_EMAIL", FROM_EMAIL);
        setIfNotNull(props, "PHONE", PHONE);
        setIfNotNull(props, "SMTP_SERVER", SMTP_SERVER);
        props.setProperty("SMTP_PORT", String.valueOf(SMTP_PORT));
        setIfNotNull(props, "SMTP_LOGIN", SMTP_LOGIN);
        setIfNotNull(props, "SMTP_PASSWORD", SMTP_PASSWORD);
        setIfNotNull(props, "COPY_EMAIL", COPY_EMAIL);

        String configPath = getPathToSetting();

        try (FileOutputStream out = new FileOutputStream(configPath)) {
            props.store(out, "Window settings");
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    private static String getPathToSetting(){
        return getPathToFolderSetting() + "/settings.properties";
    }

    private static String getPathToFolderSetting(){
        Path folderPath = Path.of(System.getProperty("user.home"), "LekoMailer");
        if (!Files.exists(folderPath)) {
            try {
                Files.createDirectories(folderPath);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return folderPath.toString();
    }

    private static void setIfNotNull(Properties props, String key, String value) {
        if (value != null) {
            props.setProperty(key, value);
        }
    }

    private static void loadSettings() {

        Properties props = new Properties();
        String configPath = getPathToSetting();

        try {
            Reader r = new InputStreamReader(new FileInputStream(configPath), StandardCharsets.UTF_8);
            props.load(r);
            SENDER = props.getProperty("SENDER", "");
            COMPANY = props.getProperty("COMPANY", "");
            FROM_EMAIL = props.getProperty("FROM_EMAIL", "");
            PHONE = props.getProperty("PHONE", "");
            SMTP_SERVER = props.getProperty("SMTP_SERVER", "");
            SMTP_PORT = Integer.parseInt(props.getProperty("SMTP_PORT", ""));
            SMTP_LOGIN = props.getProperty("SMTP_LOGIN", "");
            SMTP_PASSWORD = props.getProperty("SMTP_PASSWORD", "");
            COPY_EMAIL = props.getProperty("COPY_EMAIL", "");
        }  catch (IOException e) {
            //Ничего страшного. Нет файла с настройками.( как минимум при первом запуске)
        }

    }

    @NotNull
    private static Path getFileProgramData() {
        Path sharedPath = Path.of(
                System.getenv("ProgramData"),
                "LekoMailer",
                "stats.json"
        );

        try {
            // 1 Создаём папку если её нет
            Files.createDirectories(sharedPath.getParent());

            // 2 Создаём файл если его нет
            if (!Files.exists(sharedPath)) {
                Files.createFile(sharedPath);
                System.out.println("Файл создан");
            } else {
                System.out.println("Файл уже существует");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        return sharedPath;
    }

    private static TitledBorder titled(String title) {
        return BorderFactory.createTitledBorder(title);
    }

    // ---------- MAIN ----------
    public static void main(String[] argv) {
        Map<String, String> args = parseArgs(argv);
        boolean needGui = args.containsKey("--gui") || !(args.containsKey("--contacts") && args.containsKey("--pivot"));
        if (needGui) {
            runGui();
        } else {
            // Из аргументов CLI: можно переопределить дефолты
            if (args.containsKey("--subject")) SUBJECT_DEFAULT = args.get("--subject");
            if (args.containsKey("--smtp_server")) SMTP_SERVER = args.get("--smtp_server");
            if (args.containsKey("--smtp_port")) SMTP_PORT = Integer.parseInt(args.get("--smtp_port"));
            if (args.containsKey("--smtp_login")) SMTP_LOGIN = args.get("--smtp_login");
            if (args.containsKey("--from_email")) FROM_EMAIL = args.get("--from_email");
            runCli(args);
        }
    }
}
