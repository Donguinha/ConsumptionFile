using System.Text;
using ClosedXML.Excel;

namespace ConsoleApp1;

public abstract class Program
{
    private static readonly string Cnpj = "07652226000116";

    public static void Main(string[] args)
    {
        var xlsxFile = "C:/code/ArquivoDeConsumo/arquivo_de_consumo_mar.xlsx";
        var workbook = new XLWorkbook(xlsxFile);
        var worksheet = workbook.Worksheets.Worksheet(1);
        using var reader = new StreamReader(xlsxFile);

        var xlsxFile2 = "C:/code/ArquivoDeConsumo/Parcelado APP Folha 03 2024 contratos recusados pela sequencia.xlsx";
        var workbook2 = new XLWorkbook(xlsxFile2);
        var worksheet2 = workbook2.Worksheets.Worksheet(1);
        using var reader2 = new StreamReader(xlsxFile2);

        var txtFile = "C:/code/ArquivoDeConsumo/arquivo_de_consumo_mar_2envio.txt";
        using var writer = new StreamWriter(txtFile);

        var docList = new List<string>(11);

        writer.WriteLine($"0{DateTime.Now:yyyyMM}{Cnpj}CONSIGSIAPE".PadRight(553));

        var count = 0;

        foreach (var row in worksheet.RowsUsed())
        {
            foreach (var row2 in worksheet2.RowsUsed())
            {

                var cpf = row.Cell(8).GetValue<string>();
                var cpf2 = row2.Cell(2).GetValue<string>().Replace(".", "").Replace("-", "").TrimStart('0');

                if (cpf == cpf2 && !docList.Contains(cpf))
                {
                    docList.Add(cpf);

                    var stringBuilder = new StringBuilder()
                        .Append('1') //TIPO DO REGISTRO
                        .Append(row.Cell(2).GetValue<string>()); //ORGAO SIAPE

                    if (row.Cell(9).GetValue<string>() == "PENS")
                    {
                        stringBuilder
                            .Append(row.Cell(16).GetValue<string>().PadLeft(7, '0')) // NUMERO DA MATRICULA
                            .Append(row.Cell(7).GetValue<string>().PadLeft(8, '0')); // MATRICULA BENEFICIARIO
                    }
                    else
                    {
                        stringBuilder
                            .Append(row.Cell(7).GetValue<string>().PadLeft(7, '0')) // NUMERO DA MATRICULA
                            .Append("00000000"); // MATRICULA BENEFICIARIO
                    }

                    stringBuilder
                        .Append('4') //COMANDO
                        .Append(row2.Cell(1).GetValue<string>().PadLeft(20, '0')) // NUMERO DO CONTRATO
                        .Append("35016") // RUBRICA
                        .Append(row2.Cell(10).GetValue<string>()) // SEQUENCIA
                        .Append(FormatValue(row2.Cell(6).GetValue<string>(), 9, 2)) // VALOR
                        .Append(row2.Cell(7).GetValue<string>().PadLeft(3, '0')) // PRAZO
                        .Append(FormatValue(row2.Cell(4).GetValue<string>(), 9, 2)) // VALOR BRUTO DO CONTRATO
                        .Append(FormatValue(row2.Cell(5).GetValue<string>(), 9, 2)) // VALOR LIQUIDO A SER CREDITADO
                        .Append(FormatValue(row2.Cell(3).GetValue<string>(), 5, 2)) // IOF
                        .Append(FormatValue(row2.Cell(8).GetValue<string>(), 5, 2)) // TAXA DE JUROS MENSAL
                        .Append(FormatValue(row2.Cell(9).GetValue<string>(), 5, 2)) // CET 
                        .Append("".PadLeft(8, '0') + "".PadLeft(180, ' ') + "".PadLeft(42, '0') + "".PadLeft(181, ' '));

                    writer.WriteLine(stringBuilder);

                    count++;
                }
            }
        }

        writer.WriteLine($"9{count:D7}".PadRight(553));
    }

    private static string FormatValue(string value, int left, int right)
    {
        var valueArray = value.Split(',');

        return valueArray.Length < 2
            ? $"{valueArray[0].PadLeft(left, '0')}{"".PadLeft(right, '0')}"
            : valueArray[1].Length > 2
                ? $"{valueArray[0].PadLeft(left, '0')}{valueArray[1][..2]}"
                : $"{valueArray[0].PadLeft(left, '0')}{valueArray[1].PadRight(right, '0')}"; 
    }
}


//IMPLEMENTAÇÃO FEITA EM PRODUÇÃO

//public async Task<ValueResult<(byte[] Bytes, List<string> Errors)>> CreateTextFileAsync(Stream userFile, Stream installmentFile, CancellationToken ct)
//        {
//            var msWriter = new MemoryStream();
//            await using var writer = new StreamWriter(msWriter);
//
//            var userFileBytes = userFile.ToByteArray();
//
//            var count = 0;
//            var documentsWithError = new List<string>(11);
//
//            await writer.WriteLineAsync($"0{DateTime.Now:yyyyMM}{_cnpj}CONSIGSIAPE".PadRight(553));
//
//            using var readerInstallment = new StreamReader(installmentFile);
//
//            while (!readerInstallment.EndOfStream)
//            {
//                if (await readerInstallment.ReadLineAsync()
//                        is var lineInstallmentResult && string.IsNullOrWhiteSpace(lineInstallmentResult))
//                    return ValueResult<(byte[], List<string>)>.Failure("Empty line");
//
//                var fieldsInstallmentFile = lineInstallmentResult.Split(';');
//
//                bool documentFound = default;
//
//                var originalUserStream = new MemoryStream(userFileBytes);
//                using var disposableUserStream = new MemoryStream();
//
//                await originalUserStream.CopyToAsync(disposableUserStream, ct);
//                disposableUserStream.Seek(0, SeekOrigin.Begin);
//
//                using var readerUserData = new StreamReader(disposableUserStream);
//
//                while (!readerUserData.EndOfStream)
//                {
//                    if (await readerUserData.ReadLineAsync()
//                            is var lineUserResult && string.IsNullOrWhiteSpace(lineUserResult))
//                        return ValueResult<(byte[], List<string>)>.Failure("Empty line");
//
//                    var fieldsUserData = lineUserResult.Split(';');
//
//                    if (string.IsNullOrWhiteSpace(fieldsInstallmentFile[12]))
//                        break;
//
//                    if (UtilsExtensions.FormatDocument(fieldsInstallmentFile[12]).Equals(fieldsUserData[7]))
//                    {
//                        documentFound = true;
//
//                        var stringBuilder = new StringBuilder()
//                            .Append('1')
//                            .Append(fieldsUserData[1]);
//
//                        if (fieldsUserData[8] == "PENS")
//                        {
//                            stringBuilder
//                                .Append(fieldsUserData[15].PadLeft(7, '0'))
//                                .Append(fieldsUserData[6].PadLeft(8, '0'));
//                        }
//                        else
//                        {
//                            stringBuilder
//                                .Append(fieldsUserData[6].PadLeft(7, '0'))
//                                .Append("00000000");
//                        }
//
//                        stringBuilder
//                            .Append('4')
//                            .Append(fieldsInstallmentFile[9].PadLeft(20, '0'))
//                            .Append("35016")
//                            .Append(fieldsInstallmentFile[36])
//                            .Append(FormatValue(fieldsInstallmentFile[20], 9, 2))
//                            .Append(fieldsInstallmentFile[21].PadLeft(3, '0'))
//                            .Append(FormatValue(fieldsInstallmentFile[16], 9, 2))
//                            .Append(FormatValue(fieldsInstallmentFile[19], 9, 2))
//                            .Append(FormatValue(fieldsInstallmentFile[15], 5, 2))
//                            .Append(FormatValue(fieldsInstallmentFile[22], 5, 2))
//                            .Append(FormatValue(fieldsInstallmentFile[24], 5, 2))
//                            .Append("".PadLeft(8, '0') + "".PadLeft(180, ' ') + "".PadLeft(42, '0') + "".PadLeft(181, ' '));
//
//                        if (stringBuilder.Length == 516)
//                        {
//                            await writer.WriteLineAsync(stringBuilder, ct);
//
//                            count++;
//                        }
//                        else
//                            documentsWithError.Add(UtilsExtensions.FormatDocument(fieldsInstallmentFile[12]));
//
//                        break;
//                    }
//                }
//
//                if (!documentFound && !string.IsNullOrWhiteSpace(fieldsInstallmentFile[12]))
//                    documentsWithError.Add(UtilsExtensions.FormatDocument(fieldsInstallmentFile[12]));
//            }
//
//            await writer.WriteLineAsync($"9{count:D7}".PadRight(553));
//
//            await writer.FlushAsync();
//            msWriter.Seek(0, SeekOrigin.Begin);
//
//            return ValueResult<(byte[], List<string>)>.Success((msWriter.ToByteArray(), documentsWithError));
//        }
