// === Job Quartz que envia os e-mails ===

import java.io.File;
import java.math.BigDecimal;
import java.nio.file.Paths;
import java.text.NumberFormat;
import java.util.List;
import java.util.Locale;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.quartz.Job;
import org.quartz.JobDataMap;
import org.quartz.JobExecutionContext;
import org.quartz.JobExecutionException;

public class SendEmailsJob implements Job {

	// SMTP Zoho
	private static final String SMTP_HOST = "smtp.zoho.com";
	private static final int SMTP_PORT = 587; // TLS

	// IMAP (para copiar para itens enviados)
	private static final String IMAP_HOST = "imap.zoho.com";
	private static final int IMAP_PORT = 993;

	// Assunto do e-mail
	private static final String SUBJECT = "Boleto EasyCall";

	// Formatação de moeda (pt-BR)
	private static final NumberFormat currencyFormat = NumberFormat.getCurrencyInstance(new Locale("pt", "BR"));

	@Override
	public void execute(JobExecutionContext context) throws JobExecutionException {
		
		JobDataMap dataMap = context.getMergedJobDataMap();
		List<ClientInvoice> clients = (List<ClientInvoice>) dataMap.get("clients");
		String smtpUser = dataMap.getString("smtpUser");
		String smtpPass = dataMap.getString("smtpPass");
		String imapUser = dataMap.getString("imapUser");
		String imapPass = dataMap.getString("imapPass");
		try {
			sendBatch(clients, smtpUser, smtpPass, imapUser, imapPass);
		} catch (Exception e) {
			throw new JobExecutionException(e);
		}
	}

	private void sendBatch(List<ClientInvoice> clients, String smtpUser, String smtpPass, String imapUser, String imapPass) throws Exception {
		Properties props = new Properties();
		props.put("mail.smtp.host", SMTP_HOST);
		props.put("mail.smtp.port", String.valueOf(SMTP_PORT));
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.starttls.enable", "true");

		Session session = Session.getInstance(props, new Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(smtpUser, smtpPass);
			}
		});
		session.setDebug(false);

		for (ClientInvoice c : clients) {
			try {
				MimeMessage message = buildMessage(session, smtpUser, c);
				// Envia via SMTP
				Transport.send(message);

				// Copiar para Itens Enviados via IMAP
				try {
					appendToSentFolder(message, imapUser, imapPass);
				} catch (Exception imapEx) {
					System.err.println("Falha ao copiar mensagem para Itens Enviados via IMAP: " + imapEx.getMessage());
				}

				// Atualizar Excel marcando como Enviado
				try {
					boolean ok = ExcelReader.markRowAsSent(c.sourceExcelPath, c.excelRowIndex);
					if (ok) {
						System.out.println("Marcado como Enviado no Excel: " + c.sourceExcelPath + " (linha=" + c.excelRowIndex + ")");
					} else {
						System.err.println("Não foi possível marcar como Enviado no Excel: " + c.sourceExcelPath + " (linha=" + c.excelRowIndex + ")");
					}
				} catch (Exception ex) {
					System.err.println("Erro ao atualizar Excel para " + c.email + ": " + ex.getMessage());
					ex.printStackTrace();
				}

				System.out.println("Enviado para: " + c.email + " | CNPJ: " + c.cnpj + " | Total: " + currencyFormat.format(c.totalFinal));
			} catch (Exception e) {
				System.err.println("Erro ao enviar para " + c.email + ": " + e.getMessage());
				e.printStackTrace();
			}
		}
	}

	private MimeMessage buildMessage(Session session, String from, ClientInvoice c) throws Exception {
		MimeMessage message = new MimeMessage(session);
		message.setFrom(new InternetAddress(from));
		message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(c.email, false));
		message.setSubject(SUBJECT, "UTF-8");

		// Corpo do e-mail - texto simples com tabulação
		StringBuilder sb = new StringBuilder();
		sb.append("Prezado cliente,").append("\n\n");
		sb.append(c.cnpj).append(" - ").append(nullSafe(c.nome)).append("\n");
		// Período fixo conforme exemplo: 08/2025 (o usuário pode alterar)
		sb.append("Segue em anexo o boleto referente ao período 08/2025. Abaixo, detalhamos a composição do valor:").append("\n\n");
		sb.append(String.format("%-30s\t%-12s\t%-18s\t%s\n", "Descrição", "Quantidade", "Valor unitário", "Total"));

		// Mensalidade (sem quantidade)
		sb.append(String.format("%-30s\t%-12s\t%-18s\t%s\n", "Mensalidade", "", "", currencyFormat.format(c.mensalidade != null ? c.mensalidade : BigDecimal.ZERO)));

		// Consultas
		if (c.qtdConsultas > 0) {
			sb.append(String.format("%-30s\t%-12d\t%-18s\t%s\n", "Consultas Boa Vista", c.qtdConsultas, formatCurrency(c.valUnitConsulta), formatCurrency(c.totalConsultas)));
		}

		// Serasa (nova linha)
		if (c.qtdSerasa > 0) {
			sb.append(String.format("%-30s\t%-12d\t%-18s\t%s\n", "Serasa", c.qtdSerasa, formatCurrency(c.valUnitSerasa), formatCurrency(c.totalSerasa)));
		}

		if (c.qtdNeg > 0) {
			sb.append(String.format("%-30s\t%-12d\t%-18s\t%s\n", "Negativação", c.qtdNeg, formatCurrency(c.valUnitNeg), formatCurrency(c.totalNeg)));
		}

		if (c.qtdExc > 0) {
			sb.append(String.format("%-30s\t%-12d\t%-18s\t%s\n", "Exclusão de Negativação", c.qtdExc, formatCurrency(c.valUnitExc), formatCurrency(c.totalExc)));
		}

		if (c.qtdSms > 0) {
			sb.append(String.format("%-30s\t%-12d\t%-18s\t%s\n", "SMS", c.qtdSms, formatCurrency(c.valUnitSms), formatCurrency(c.totalSms)));
		}

		// NF-e (Nota Fiscal Eletrônica)
		if (c.qtdNf > 0) {
			sb.append(String.format("%-30s\t%-12d\t%-18s\t%s\n", "NF-e", c.qtdNf, formatCurrency(c.valUnitNf), formatCurrency(c.totalNf)));
		}

		sb.append("\t\tTOTAL\t").append(formatCurrency(c.totalFinal)).append("\n\n");
		sb.append("Em caso de dúvidas, estamos à disposição.\n\n");
		sb.append("Atenciosamente,\n\n");
		sb.append("EasyCall\n");

		// Criar multipart (texto + anexo)
		MimeBodyPart textPart = new MimeBodyPart();
		textPart.setText(sb.toString(), "UTF-8");

		Multipart multipart = new MimeMultipart();
		multipart.addBodyPart(textPart);

		// Anexo
		if (c.arquivo != null && !c.arquivo.trim().isEmpty()) {
			File f = new File(c.arquivo);
			if (!f.exists()) {
				// tentar interpretar caminho relativo ao Excel
				String alt = Paths.get(c.arquivo).toAbsolutePath().toString();
				System.err.println("Arquivo anexo não encontrado: " + c.arquivo + " | tentando absoluto: " + alt);
			} else {
				MimeBodyPart attachPart = new MimeBodyPart();
				DataSource source = new FileDataSource(f);
				attachPart.setDataHandler(new DataHandler(source));
				attachPart.setFileName(f.getName());
				multipart.addBodyPart(attachPart);
			}
		} else {
			System.out.println("Aviso: nenhum arquivo especificado para " + c.email + " (CNPJ: " + c.cnpj + ")");
		}

		message.setContent(multipart);
		message.saveChanges();
		return message;
	}

	private void appendToSentFolder(MimeMessage message, String imapUser, String imapPass) throws Exception {
		Properties props = new Properties();
		props.put("mail.imap.ssl.enable", "true");
		props.put("mail.imap.host", IMAP_HOST);
		props.put("mail.imap.port", String.valueOf(IMAP_PORT));

		Session imapSession = Session.getInstance(props);
		Store store = imapSession.getStore("imap");
		store.connect(IMAP_HOST, imapUser, imapPass);

		// Nome da pasta Sent em PT-BR pode variar; tentamos pastas comuns
		String[] candidates = new String[] { "Sent", "Sent Items", "Itens Enviados", "Itens Enviados (Padrão)", "Enviados", "INBOX.Sent" };
		Folder sent = null;
		for (String name : candidates) {
			try {
				Folder f = store.getFolder(name);
				if (f != null && (f.exists() || f.getName() != null)) {
					sent = f;
					break;
				}
			} catch (Exception ignored) {
			}
		}
		if (sent == null) {
			// fallback para pasta padrão "Sent"
			sent = store.getFolder("Sent");
		}
		if (!sent.exists()) {
			// tenta criar
			try {
				sent.create(Folder.HOLDS_MESSAGES);
			} catch (Exception ignore) {
			}
		}

		sent.open(Folder.READ_WRITE);
		// Observação: ao usar appendMessages, a mensagem deve pertencer à mesma Session ou ser convertida;
		// como temos a MimeMessage, podemos usar: sent.appendMessages(new Message[]{message});
		sent.appendMessages(new Message[] { message });
		sent.close(false);
		store.close();
	}
	
	public void executeNow(List<ClientInvoice> clients, String smtpUser, String smtpPass, String imapUser, String imapPass) {
        try {
            System.out.println("Enviando e-mails diretamente (sem agendamento)...");
            for (ClientInvoice c : clients) {
                try {
                    System.out.println("📧 Enviando e-mail para: " + c.email + " (" + c.nome + ")");
                    EmailService.sendEmailWithAttachment(smtpUser, smtpPass, imapUser, imapPass, c);
                    // marcar no Excel como enviado
                    try {
                        boolean ok = ExcelReader.markRowAsSent(c.sourceExcelPath, c.excelRowIndex);
                        if (ok) {
                            System.out.println("Marcado como Enviado no Excel: " + c.sourceExcelPath + " (linha=" + c.excelRowIndex + ")");
                        } else {
                            System.err.println("Não foi possível marcar como Enviado no Excel: " + c.sourceExcelPath + " (linha=" + c.excelRowIndex + ")");
                        }
                    } catch (Exception ex) {
                        System.err.println("Erro ao atualizar Excel para " + c.email + ": " + ex.getMessage());
                        ex.printStackTrace();
                    }
                } catch (Exception e) {
                    System.err.println("Erro ao enviar para " + c.email + ": " + e.getMessage());
                    e.printStackTrace();
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

	private String nullSafe(String s) {
		return s == null ? "" : s;
	}

	private String formatCurrency(BigDecimal v) {
		return v == null ? currencyFormat.format(BigDecimal.ZERO) : currencyFormat.format(v);
	}
}