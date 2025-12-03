import javax.naming.*;
import javax.naming.directory.*;
import java.io.*;
import java.net.InetAddress;
import java.util.*;
import java.util.concurrent.*;

public class AutoADScannerLoot {

    private static final int THREADS = 50;
    private static volatile int progress = 0;

    // patterns to search inside shares
    private static final String[] LOOT_PATTERNS = {
            "password", "pass", "cred", "credential",
            "secret", "vpn", "key", ".pem", ".pfx",
            ".kdbx", ".json", ".config", ".rdp", ".txt"
    };

    public static void main(String[] args) throws Exception {

        System.out.println("[*] Detecting domain...");

        String domain = System.getenv("USERDNSDOMAIN");
        if (domain == null) {
            System.out.println("[!] Not in domain.");
            return;
        }

        String ldapURL = "ldap://" + domain;
        System.out.println("[+] Domain   : " + domain);
        System.out.println("[+] LDAP URL : " + ldapURL);

        System.out.println("[*] Fetching computers...");
        List<String> computers = getADComputers(ldapURL);
        System.out.println("[+] Computers: " + computers.size());

        ExecutorService resolvePool = Executors.newFixedThreadPool(THREADS);

        // --------------------------------------------------------
        // STEP 1: Resolve IPs
        // --------------------------------------------------------
        List<String> ips = Collections.synchronizedList(new ArrayList<>());
        progress = 0;
        System.out.println("[*] Resolving IPs...");

        for (String host : computers) {
            resolvePool.submit(() -> {
                String ip = resolveIP(host);
                if (ip != null) ips.add(ip);
                updateProgress("Resolving", computers.size());
            });
        }

        resolvePool.shutdown();
        resolvePool.awaitTermination(10, TimeUnit.MINUTES);

        System.out.println("\n[+] Alive hosts: " + ips.size());

        // --------------------------------------------------------
        // STEP 2: Scan shares + search loot
        // --------------------------------------------------------
        ExecutorService smbPool = Executors.newFixedThreadPool(THREADS);

        FileWriter report = new FileWriter("report.txt");
        FileWriter loot = new FileWriter("loot.txt");

        report.write("=== SMB Share Report ===\n\n");
        loot.write("=== LOOT FILES FOUND ===\n\n");

        Object lock = new Object();

        progress = 0;
        System.out.println("[*] Scanning SMB & searching loot...");

        for (String ip : ips) {
            smbPool.submit(() -> {
                List<String> shares = getShares(ip);

                synchronized (lock) {
                    try {
                        report.write("IP: " + ip + "\n");
                        if (shares.isEmpty()) {
                            report.write("   No shares / access denied\n\n");
                        } else {
                            for (String s : shares)
                                report.write("   " + s + "\n");
                            report.write("\n");
                        }
                    } catch (Exception ignored) {}
                }

                // LOOT SCAN
                for (String share : shares) {
                    String shareName = share.split(" ")[0];

                    List<String> found = searchLoot(ip, shareName);

                    if (!found.isEmpty()) {
                        synchronized (lock) {
                            try {
                                loot.write("Host: " + ip + " Share: " + shareName + "\n");
                                for (String f : found)
                                    loot.write("   " + f + "\n");
                                loot.write("\n");
                            } catch (Exception ignored) {}
                        }
                    }
                }

                updateProgress("Scanning", ips.size());
            });
        }

        smbPool.shutdown();
        smbPool.awaitTermination(20, TimeUnit.MINUTES);

        report.close();
        loot.close();

        System.out.println("\n[+] Saved report.txt and loot.txt");
    }

    // ----------------------------------------------------------
    // LDAP
    // ----------------------------------------------------------
    private static List<String> getADComputers(String url) throws Exception {

        List<String> list = new ArrayList<>();

        Hashtable<String,String> env = new Hashtable<>();
        env.put(Context.INITIAL_CONTEXT_FACTORY,"com.sun.jndi.ldap.LdapCtxFactory");
        env.put(Context.PROVIDER_URL, url);
        env.put(Context.SECURITY_AUTHENTICATION, "GSSAPI");

        DirContext ctx = new InitialDirContext(env);

        Attributes root = ctx.getAttributes("", new String[]{"defaultNamingContext"});
        String baseDN = root.get("defaultNamingContext").get().toString();

        SearchControls sc = new SearchControls();
        sc.setSearchScope(SearchControls.SUBTREE_SCOPE);
        sc.setReturningAttributes(new String[]{"dNSHostName","sAMAccountName"});

        NamingEnumeration<SearchResult> answer = ctx.search(baseDN,"(objectClass=computer)", sc);

        while (answer.hasMore()) {
            SearchResult sr = answer.next();
            Attributes a = sr.getAttributes();

            String host = null;

            if (a.get("dNSHostName") != null)
                host = a.get("dNSHostName").get().toString();
            else if (a.get("sAMAccountName") != null)
                host = a.get("sAMAccountName").get().toString().replace("$","");

            if (host != null) list.add(host);
        }

        return list;
    }

    // ----------------------------------------------------------
    private static String resolveIP(String host) {
        try {
            InetAddress inet = InetAddress.getByName(host);
            return inet.getHostAddress();
        } catch (Exception e) { return null; }
    }

    // ----------------------------------------------------------
    private static List<String> getShares(String ip) {
        List<String> shares = new ArrayList<>();
        try {
            Process p = new ProcessBuilder("cmd.exe","/c","net view \\\\" + ip)
                    .redirectErrorStream(true).start();

            BufferedReader br = new BufferedReader(new InputStreamReader(p.getInputStream()));
            String line;

            while ((line = br.readLine()) != null) {
                if (line.contains("Disk") || line.contains("IPC$"))
                    shares.add(line.trim());
            }
        } catch (Exception ignored) {}
        return shares;
    }

    // ----------------------------------------------------------
    // Search files using DIR /S /B
    // ----------------------------------------------------------
    private static List<String> searchLoot(String ip, String share) {
        List<String> found = new ArrayList<>();

        try {
            // path example: \\10.0.0.5\Public
            String path = "\\\\" + ip + "\\" + share;

            Process p = new ProcessBuilder("cmd.exe","/c","dir \"" + path + "\" /s /b")
                    .redirectErrorStream(true)
                    .start();

            BufferedReader br = new BufferedReader(new InputStreamReader(p.getInputStream()));
            String line;

            while ((line = br.readLine()) != null) {
                String lower = line.toLowerCase();

                for (String pattern : LOOT_PATTERNS)
                    if (lower.contains(pattern)) found.add(line);
            }

        } catch (Exception ignored) {}

        return found;
    }

    // ----------------------------------------------------------
    private static synchronized void updateProgress(String phase, int total) {
        progress++;
        int width = 40;
        double pct = (double) progress / total;
        int filled = (int)(pct * width);

        StringBuilder sb = new StringBuilder();
        sb.append("\r").append(phase).append(" [");
        for (int i = 0; i < filled; i++) sb.append("=");
        for (int i = filled; i < width; i++) sb.append(" ");
        sb.append("] ").append((int)(pct*100))
          .append("% ").append(progress).append("/").append(total);

        System.out.print(sb.toString());
    }
}
