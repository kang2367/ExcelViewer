using System.Globalization;
using System.Text;

namespace ExcelDropViewer
{
    internal static class DigiKeySearchResultFormatter
    {
        public static string Format(DigiKeyProductSummary summary)
        {
            var builder = new StringBuilder();
            builder.AppendLine("[Digi-Key 부품 검색 결과 - Cut Tape(CT) 기준]");
            builder.AppendLine("----------------------------------------");
            builder.AppendLine($"- 검색 부품 번호 : {summary.SearchedPartNumber}");
            builder.AppendLine($"- 검색 제조사 명 : {summary.SearchedManufacturer}");
            builder.AppendLine($"- 디지키 파트 번호 : {summary.DigiKeyPartNumber}");
            builder.AppendLine($"- 디지키 등록 제조사 : {summary.DigiKeyManufacturer}");
            builder.AppendLine($"- 포장 형태 : {summary.PackagingType}");
            builder.AppendLine($"- 재고 수량 : {summary.QuantityAvailable:N0} 개");
            builder.AppendLine("- 컷 테이프 수량별 단가 (KRW):");

            if (summary.PriceTiers.Count == 0)
            {
                builder.AppendLine("  * 가격 정보 없음");
            }
            else
            {
                foreach (var tier in summary.PriceTiers)
                {
                    builder.AppendLine(
                        $"  * {tier.BreakQuantity.ToString("#,##0", CultureInfo.GetCultureInfo("ko-KR"))}개 이상 : {tier.FormattedUnitPrice}");
                }
            }

            builder.Append("----------------------------------------");
            return builder.ToString();
        }
    }
}
