import com.poiji.bind.Poiji;
import com.poiji.option.PoijiExcelType;
import org.springframework.web.multipart.MultipartFile;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.util.Collections;
import java.util.List;

public List<ExcelDataModel> processExcel(MultipartFile file) {
    try (InputStream inputStream = file.getInputStream();
         PushbackInputStream pushbackInputStream = new PushbackInputStream(inputStream, 8)) {

        // Read the first few bytes to identify the file type
        byte[] header = new byte[8];
        pushbackInputStream.read(header);
        pushbackInputStream.unread(header); // Push back bytes for Poiji to read

        PoijiExcelType excelType;
        if (isXLS(header)) {
            excelType = PoijiExcelType.XLS;
        } else if (isXLSX(header)) {
            excelType = PoijiExcelType.XLSX;
        } else {
            throw new IllegalArgumentException("Unsupported file type. Please upload an Excel file.");
        }

        return Poiji.fromExcel(pushbackInputStream, excelType, ExcelDataModel.class);
    } catch (Exception e) {
        e.printStackTrace();
        return Collections.emptyList(); // Returns empty list on error
    }
}

private boolean isXLS(byte[] header) {
    return header[0] == (byte) 0xD0 && header[1] == (byte) 0xCF && header[2] == (byte) 0x11 && header[3] == (byte) 0xE0;
}

private boolean isXLSX(byte[] header) {
    return header[0] == (byte) 0x50 && header[1] == (byte) 0x4B && header[2] == (byte) 0x03 && header[3] == (byte) 0x04;
}
