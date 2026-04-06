import java.io.File;
import java.math.BigDecimal;
import java.nio.file.Paths;
import java.text.NumberFormat;
import java.util.Locale;
import java.util.Properties;
import java.util.concurrent.ThreadLocalRandom;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

public class EmailService {

	private static final String SMTP_HOST = "smtp.zoho.com";
	private static final int SMTP_PORT = 587;
	private static final String PERIORO = "03/2026";
	

	private static final NumberFormat currencyFormat = NumberFormat.getCurrencyInstance(new Locale("pt", "BR"));

	public static void sendEmailWithAttachment(String smtpUser, String smtpPass, String imapUser, String imapPass, ClientInvoice c) throws Exception {
		
		// Random entre 45 e 99 segundos por conta do último bloqueio
		int delay = ThreadLocalRandom.current().nextInt(45000, 99000); 
		System.out.println("⏳ Aguardando " + (delay / 1000.0) + " segundos antes do próximo envio...");
		Thread.sleep(delay);

		// --- Config SMTP ---
		Properties props = new Properties();
		props.put("mail.smtp.host", SMTP_HOST);
		props.put("mail.smtp.port", String.valueOf(SMTP_PORT));
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.starttls.enable", "true");

		props.put("mail.smtp.ssl.trust", "*"); // Confia em todos os certificados
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.starttls.required", "true");

		Session session = Session.getInstance(props, new Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(smtpUser, smtpPass);
			}
		});
		session.setDebug(false);

		// --- Construir mensagem ---
		MimeMessage message = new MimeMessage(session);
		message.setFrom(new InternetAddress(smtpUser));
		if (c.email == null || c.email.trim().isEmpty()) {
			throw new IllegalArgumentException("Cliente sem e-mail: CNPJ=" + c.cnpj + " Nome=" + c.nome);
		}
		message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(c.email, false));
		message.setSubject("Boleto EasyCall", "UTF-8");

		StringBuilder sb = new StringBuilder();

		sb.append("<!DOCTYPE html>");
		sb.append("<html lang='pt-BR'>");
		sb.append("<body style='font-family: Arial, sans-serif; color: #333;'>");
		sb.append("<p>Prezado cliente,</p>");
		sb.append("<p>")
		  .append((c.cnpj != null ? c.cnpj : ""))
		  .append(" - <strong>")
		  .append(nullSafe(c.nome))
		  .append("</strong></p>");
		sb.append("<p>Segue em anexo o boleto referente ao período "+PERIORO+".<br>");
		sb.append("Abaixo, detalhamos a composição do valor:</p>");

		sb.append("<div style='max-width: 600px; margin-left: 0;'>"); // Alinhada à esquerda
		sb.append("<table style='border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 14px;'>");
		sb.append("<thead>");
		sb.append("<tr style='background-color: #f2f2f2;'>");
		sb.append("<th style='text-align: left; padding: 8px; border: 1px solid #ddd;'>Descrição</th>");
		sb.append("<th style='text-align: center; padding: 8px; border: 1px solid #ddd;'>Quantidade</th>");
		sb.append("<th style='text-align: right; padding: 8px; border: 1px solid #ddd;'>Valor unitário</th>");
		sb.append("<th style='text-align: right; padding: 8px; border: 1px solid #ddd;'>Total</th>");
		sb.append("</tr>");
		sb.append("</thead>");
		sb.append("<tbody>");

		sb.append(String.format(
		    "<tr><td style='padding:8px;border:1px solid #ddd;'>Mensalidade</td><td style='text-align:center;padding:8px;border:1px solid #ddd;'>-</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>-</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>%s</td></tr>",
		    formatCurrency(c.mensalidade)));

		sb.append(String.format(
		    "<tr><td style='padding:8px;border:1px solid #ddd;'>Consultas Boa Vista</td><td style='text-align:center;padding:8px;border:1px solid #ddd;'>%d</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>%s</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>%s</td></tr>",
		    c.qtdConsultas, formatCurrency(c.valUnitConsulta), formatCurrency(c.totalConsultas)));

		sb.append(String.format(
		    "<tr><td style='padding:8px;border:1px solid #ddd;'>Serasa</td><td style='text-align:center;padding:8px;border:1px solid #ddd;'>%d</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>%s</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>%s</td></tr>",
		    c.qtdSerasa, formatCurrency(c.valUnitSerasa), formatCurrency(c.totalSerasa)));

		sb.append(String.format(
		    "<tr><td style='padding:8px;border:1px solid #ddd;'>Negativação</td><td style='text-align:center;padding:8px;border:1px solid #ddd;'>%d</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>%s</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>%s</td></tr>",
		    c.qtdNeg, formatCurrency(c.valUnitNeg), formatCurrency(c.totalNeg)));

		sb.append(String.format(
		    "<tr><td style='padding:8px;border:1px solid #ddd;'>Exclusão de Negativação</td><td style='text-align:center;padding:8px;border:1px solid #ddd;'>%d</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>%s</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>%s</td></tr>",
		    c.qtdExc, formatCurrency(c.valUnitExc), formatCurrency(c.totalExc)));

		sb.append(String.format(
		    "<tr><td style='padding:8px;border:1px solid #ddd;'>SMS</td><td style='text-align:center;padding:8px;border:1px solid #ddd;'>%d</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>%s</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>%s</td></tr>",
		    c.qtdSms, formatCurrency(c.valUnitSms), formatCurrency(c.totalSms)));

		sb.append(String.format(
		    "<tr style='background-color:#f9f9f9;font-weight:bold;'><td colspan='3' style='text-align:right;padding:8px;border:1px solid #ddd;'>TOTAL</td><td style='text-align:right;padding:8px;border:1px solid #ddd;'>%s</td></tr>",
		    formatCurrency(c.totalFinal)));

		sb.append("</tbody>");
		sb.append("</table>");
		sb.append("</div>");

		sb.append("<p style='margin-top:20px;'>Em caso de dúvidas, estamos à disposição.</p>");
		sb.append("<p>Atenciosamente,<br><strong>EasyCall</strong></p>");
		sb.append("</body></html>");

		MimeBodyPart textPart = new MimeBodyPart();
		textPart.setContent(sb.toString(), "text/html; charset=UTF-8");		

		Multipart multipart = new MimeMultipart();
		multipart.addBodyPart(textPart);

		// --- Anexo ---
		if (c.arquivo != null && !c.arquivo.trim().isEmpty()) {
			File f = new File(c.arquivo);
			if (!f.exists()) {
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

		// --- Enviar via SMTP ---
		Transport.send(message);

		System.out.println("Enviado: " + c.email + " | CNPJ: " + c.cnpj + " | Total: " + formatCurrency(c.totalFinal));
	}

	private static String nullSafe(String s) {
		return s == null ? "" : s;
	}

	private static String formatCurrency(BigDecimal v) {
		return v == null ? currencyFormat.format(BigDecimal.ZERO) : currencyFormat.format(v);
	}
}