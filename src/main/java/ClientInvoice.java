
import java.math.BigDecimal;

public class ClientInvoice {
        String nome;
        String cnpj;
        String email;
        int qtdConsultas;
        BigDecimal valUnitConsulta;
        BigDecimal totalConsultas;
        int qtdNeg;
        BigDecimal valUnitNeg;
        BigDecimal totalNeg;
        int qtdExc;
        BigDecimal valUnitExc;
        BigDecimal totalExc;
        int qtdSms;
        BigDecimal valUnitSms;
        BigDecimal totalSms;
        BigDecimal mensalidade;
        BigDecimal totalFinal;
        String arquivo; // caminho do boleto
        // Nova informação: linha (0-based) no Excel e caminho do arquivo Excel de origem
        int excelRowIndex;
        String sourceExcelPath;

        public ClientInvoice(String nome, String cnpj, String email,
                             int qtdConsultas, BigDecimal valUnitConsulta, BigDecimal totalConsultas,
                             int qtdNeg, BigDecimal valUnitNeg, BigDecimal totalNeg,
                             int qtdExc, BigDecimal valUnitExc, BigDecimal totalExc,
                             int qtdSms, BigDecimal valUnitSms, BigDecimal totalSms,
                             BigDecimal mensalidade, BigDecimal totalFinal, String arquivo,
                             int excelRowIndex, String sourceExcelPath) {
            this.nome = nome;
            this.cnpj = cnpj;
            this.email = email;
            this.qtdConsultas = qtdConsultas;
            this.valUnitConsulta = valUnitConsulta;
            this.totalConsultas = totalConsultas;
            this.qtdNeg = qtdNeg;
            this.valUnitNeg = valUnitNeg;
            this.totalNeg = totalNeg;
            this.qtdExc = qtdExc;
            this.valUnitExc = valUnitExc;
            this.totalExc = totalExc;
            this.qtdSms = qtdSms;
            this.valUnitSms = valUnitSms;
            this.totalSms = totalSms;
            this.mensalidade = mensalidade;
            this.totalFinal = totalFinal;
            this.arquivo = arquivo;
            this.excelRowIndex = excelRowIndex;
            this.sourceExcelPath = sourceExcelPath;
        }
    }