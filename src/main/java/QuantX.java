import com.beust.jcommander.IStringConverter;
import com.beust.jcommander.JCommander;
import com.beust.jcommander.Parameter;

import com.beust.jcommander.ParameterException;
import java.io.*;
import java.util.*;

public class QuantX {

    public static class FileListConverter implements IStringConverter<List<String>> {
        @Override
        public List<String> convert(String files) {
            String [] paths = files.split(",");
            List<String> fileList = new ArrayList<>();
            Collections.addAll(fileList, paths);

            return fileList;
        }
    }

    @Parameter(names={"--fn-guide", "-fg"}, listConverter = FileListConverter.class)
    private final List<String> fnGuides = new ArrayList<>();

    @Parameter(names={"--full-catalog", "-fc"})
    private String fullCatalogFilename;

    @Parameter(names={"--shoppinfg-guide", "-sg"})
    private String shoppingGuideFilename;

    @Parameter(names={"--output", "-o"})
    private String outputFilename;

    public QuantX() {
        this.outputFilename = "output.xlsx";
    }

    public void run() throws IOException {

        Map<String, Company> companies = new HashMap<>();

        for(String fileName : this.fnGuides) {
            FnGuide fnGuide = new FnGuide(companies);

            fnGuide.load(fileName);
        }

        FullCatalog fullCatalog = new FullCatalog(companies);

        if (this.fullCatalogFilename != null) {
            fullCatalog.load(this.fullCatalogFilename);
        }

        if (this.shoppingGuideFilename != null)  {
            ShoppingGuide shoppingGuide = new ShoppingGuide(companies);

            shoppingGuide.load(this.shoppingGuideFilename);
        }

        fullCatalog.calculate();
        //fullCatalog.save("./data.json");
        fullCatalog.saveToExcel(this.outputFilename);

    }

    public static void main(String [] args) throws IOException {
        QuantX quantX = new QuantX();

        JCommander commander = JCommander.newBuilder()
            .addObject(quantX)
            .build();

        try {
            commander.parse(args);
            quantX.run();
        } catch (ParameterException ignore) {
            commander.usage();
        }
    }

}
