package au.org.cpc.es;

import java.io.IOException;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class CreatePptxServlet extends javax.servlet.http.HttpServlet {
    static private class SlideWriter {
        private enum State {
            PAGE_BREAK,
            TITLE,
            BODY,
            CREDITS,
        }

        private XMLSlideShow show;
        private XSLFSlideLayout blankLayout;
        private XSLFSlideLayout textLayout;
        private State state = State.PAGE_BREAK;
        private XSLFSlide slide = null;
        private boolean[] placeholderUsed = new boolean[2];

        public SlideWriter(XMLSlideShow show) {
            while (!show.getSlides().isEmpty()) {
                show.removeSlide(0);
            }
            this.show = show;

            java.util.List<XSLFSlideMaster> masters = show.getSlideMasters();
            XSLFSlideMaster master = masters.get(masters.size() - 1);

            this.blankLayout = master.getLayout(SlideLayout.BLANK);
            this.textLayout = master.getLayout(SlideLayout.TEXT);
        }

        private void initSlide(XSLFSlide slide) {
            this.slide = slide;
            for (int i = 0; i < placeholderUsed.length; ++i) {
                placeholderUsed[i] = false;
            }
        }

        public void processLine(String line) {
            line = line.trim();
            if (line.isEmpty()) {
                if (state == State.PAGE_BREAK) {
                    show.createSlide(blankLayout);
                }
                state = State.PAGE_BREAK;
                return;
            }

            if (state == State.PAGE_BREAK) {
                initSlide(show.createSlide(textLayout));
                state = State.BODY;
            }
            if (line.equals("##title")) {
                state = State.TITLE;
                return;
            } else if (line.equals("##credits")) {
                state = State.CREDITS;
                addText(1, " ", 0.5);
                return;
            }

            switch (state) {
            case TITLE:
                addText(0, line, 1.0);
                state = State.BODY;
                break;
            case BODY:
                addText(1, line, 1.0);
                break;
            case CREDITS:
                addText(1, line, 0.5);
                break;
            default:
                throw new IllegalStateException();
            }
        }

        private void addText(int placeholderIndex, String line, double sizeRatio) {
            XSLFTextShape text = slide.getPlaceholder(placeholderIndex);
            if (!placeholderUsed[placeholderIndex]) {
                placeholderUsed[placeholderIndex] = true;
                text.clearText();
            }
            XSLFTextParagraph para = text.addNewTextParagraph();
            para.setBullet(false);
            para.setIndentLevel(0);
            XSLFTextRun run = para.addNewTextRun();
            run.setFontSize(run.getFontSize() * sizeRatio);
            run.setText(line);
        }
    }

    private byte[] pptxTemplate;

    @Override
    public void init() throws ServletException {
        try {
            java.io.InputStream is = getServletContext()
                .getResourceAsStream("/WEB-INF/template.pptx");

            // It's troublesome to prepare pptx files with 0 slides (from Google
            // Docs), so just remove any slides here.
            XMLSlideShow show = new XMLSlideShow(is);
            for (int n = show.getSlides().size(); --n >= 0;) {
                show.removeSlide(n);
            }
            java.io.ByteArrayOutputStream os = new java.io.ByteArrayOutputStream();
            show.write(os);

            pptxTemplate = os.toByteArray();
        } catch (IOException e) {
            throw new ServletException(e);
        }
    }

    @Override
    public void doPost(HttpServletRequest req, HttpServletResponse resp) throws IOException {
        XMLSlideShow show = new XMLSlideShow(new java.io.ByteArrayInputStream(pptxTemplate));
        SlideWriter writer = new SlideWriter(show);
        java.io.BufferedReader reader = req.getReader();
        for (;;) {
            String line = reader.readLine();
            if (line == null) break;
            writer.processLine(line);
        }

        resp.setContentType("application/octet-stream");
        show.write(resp.getOutputStream());
    }
}
