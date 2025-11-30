import java.net.ServerSocket;

public class SocketTest {
    public static void main(String[] args) throws Exception {
        int port = 3890;
        System.out.println("Trying to open port " + port + "...");
        ServerSocket server = new ServerSocket(port);
        System.out.println("SUCCESS! Listening on port " + port);
        System.out.println("Press Ctrl+C to stop");
        Thread.sleep(60000);
        server.close();
    }
}