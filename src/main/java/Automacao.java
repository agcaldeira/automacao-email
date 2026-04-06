import java.util.List;

public class Automacao {
	
	
	// >>>>> Sempre ajustar                      <<<<<
	// >>>>>  - O nome do arquivo                <<<<<
	// >>>>>  - O PERIORO na classe EmailService <<<<<

    private static final String DEFAULT_EXCEL = "C:\\Users\\agcal\\Dropbox\\EasyCall_Comercial\\2. Financeiro\\Controle de Mensalidade_cobrança mensal\\2026-04 - Mensalidade.xlsx";

    // SMTP Zoho
    private static final String ZOHO_SMTP_USER = System.getenv("ZOHO_SMTP_USER"); // seu@seudominio.com
    private static final String ZOHO_SMTP_APP_PASSWORD = System.getenv("ZOHO_SMTP_APP_PASSWORD"); // senha de app SMTP

    // IMAP (para copiar para itens enviados)
    private static final String ZOHO_IMAP_USER = System.getenv("ZOHO_IMAP_USER"); // normalmente igual ao SMTP_USER
    private static final String ZOHO_IMAP_PASSWORD = System.getenv("ZOHO_IMAP_PASSWORD"); // senha de app IMAP (pode ser a mesma do SMTP)

    public static void main(String[] args) throws Exception {
        String excelPath = (args.length > 0 && args[0] != null && !args[0].isEmpty()) ? args[0] : DEFAULT_EXCEL;

        // Checar credenciais
        if (ZOHO_SMTP_USER == null || ZOHO_SMTP_APP_PASSWORD == null || ZOHO_IMAP_USER == null || ZOHO_IMAP_PASSWORD == null) {
            System.err.println("❌ Configure as variáveis de ambiente:");
            System.err.println("ZOHO_SMTP_USER, ZOHO_SMTP_APP_PASSWORD, ZOHO_IMAP_USER e ZOHO_IMAP_PASSWORD");
            System.exit(1);
        }

        // Ler planilha e montar lista de clientes
        List<ClientInvoice> clients = ExcelReader.readClients(excelPath);

        if (clients.isEmpty()) {
            System.out.println("Nenhum cliente encontrado no arquivo Excel. Saindo.");
            return;
        }

        System.out.println("📤 Iniciando envio de " + clients.size() + " e-mails...");

        // Criar e executar job diretamente
        SendEmailsJob job = new SendEmailsJob();
        job.executeNow(clients, ZOHO_SMTP_USER, ZOHO_SMTP_APP_PASSWORD, ZOHO_IMAP_USER, ZOHO_IMAP_PASSWORD);

        System.out.println("✅ Todos os e-mails foram enviados (ou processados com erro, se houver).");
    }

}
