using System.Globalization;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelDropViewer
{
    internal sealed class DigiKeyApiService
    {
        private const string TokenUrl = "https://api.digikey.com/v1/oauth2/token";
        private const string KeywordSearchUrl = "https://api.digikey.com/products/v4/search/keyword";
        private const int CutTapePackageTypeId = 2;
        private const int TapeAndReelPackageTypeId = 1;

        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            PropertyNameCaseInsensitive = true
        };

        private readonly HttpClient _httpClient;

        public DigiKeyApiService(HttpClient? httpClient = null)
        {
            _httpClient = httpClient ?? new HttpClient();
        }

        public async Task<DigiKeyProductSummary> SearchByPartNumberAsync(
            string partNumber,
            string? manufacturer,
            DigikeyConfig config,
            CancellationToken cancellationToken = default)
        {
            var normalizedManufacturer = DigiKeyManufacturerNameNormalizer.Normalize(manufacturer);
            var accessToken = await RequestAccessTokenAsync(config, cancellationToken).ConfigureAwait(false);
            return await SearchKeywordAsync(
                    partNumber,
                    normalizedManufacturer,
                    config.ClientId,
                    accessToken,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        private async Task<string> RequestAccessTokenAsync(DigikeyConfig config, CancellationToken cancellationToken)
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, TokenUrl)
            {
                Content = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    ["client_id"] = config.ClientId,
                    ["client_secret"] = config.ClientSecret,
                    ["grant_type"] = "client_credentials"
                })
            };

            using var response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException(
                    $"Digi-Key OAuth 토큰 발급 실패 ({(int)response.StatusCode}): {ExtractApiError(body)}");
            }

            var tokenResponse = JsonSerializer.Deserialize<DigiKeyTokenResponse>(body, JsonOptions);
            if (string.IsNullOrWhiteSpace(tokenResponse?.AccessToken))
            {
                throw new InvalidOperationException("Digi-Key OAuth 응답에 access_token이 없습니다.");
            }

            return tokenResponse.AccessToken;
        }

        private async Task<DigiKeyProductSummary> SearchKeywordAsync(
            string partNumber,
            string manufacturer,
            string clientId,
            string accessToken,
            CancellationToken cancellationToken)
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, KeywordSearchUrl);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            request.Headers.Add("X-DIGIKEY-Client-Id", clientId);
            request.Headers.Add("X-DIGIKEY-Locale-Site", "KR");
            request.Headers.Add("X-DIGIKEY-Locale-Language", "ko");
            request.Headers.Add("X-DIGIKEY-Locale-Currency", "KRW");

            var payload = JsonSerializer.Serialize(new DigiKeyKeywordSearchRequest
            {
                Keywords = BuildSearchKeywords(partNumber, manufacturer),
                Limit = 25,
                Offset = 0
            });
            request.Content = new StringContent(payload, Encoding.UTF8, "application/json");

            using var response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException(
                    $"Digi-Key 부품 검색 API 호출 실패 ({(int)response.StatusCode}): {ExtractApiError(body)}");
            }

            var searchResponse = JsonSerializer.Deserialize<DigiKeyKeywordSearchResponse>(body, JsonOptions);
            var (product, variation) = SelectCutTapeProductAndVariation(searchResponse, partNumber, manufacturer);
            return MapProductSummary(partNumber, manufacturer, product, variation);
        }

        private static string BuildSearchKeywords(string partNumber, string manufacturer)
        {
            var trimmedPartNumber = partNumber.Trim();
            if (string.IsNullOrWhiteSpace(manufacturer))
            {
                return trimmedPartNumber;
            }

            return $"{trimmedPartNumber} {manufacturer.Trim()}";
        }

        private static (DigiKeyProduct Product, DigiKeyProductVariation Variation) SelectCutTapeProductAndVariation(
            DigiKeyKeywordSearchResponse? searchResponse,
            string searchedPartNumber,
            string searchedManufacturer)
        {
            var products = CollectSearchProducts(searchResponse);
            if (products.Count == 0)
            {
                throw new InvalidOperationException($"Digi-Key에서 '{searchedPartNumber}'에 대한 검색 결과가 없습니다.");
            }

            var matches = new List<(DigiKeyProduct Product, DigiKeyProductVariation Variation, int Priority)>();
            foreach (var product in products)
            {
                if (IsExcludedReelOnlyProduct(product))
                {
                    continue;
                }

                var variation = FindCutTapeVariation(product, searchedPartNumber);
                if (variation != null)
                {
                    matches.Add((
                        product,
                        variation,
                        ScoreCutTapeMatch(product, variation, searchedPartNumber, searchedManufacturer)));
                }
            }

            if (matches.Count == 0)
            {
                throw new InvalidOperationException(
                    $"Digi-Key에서 '{searchedPartNumber}'의 Cut Tape(CT) 패키징을 찾을 수 없습니다. " +
                    "Tape & Reel(TR) 항목은 제외되었습니다.");
            }

            var bestMatch = matches
                .OrderByDescending(match => match.Priority)
                .ThenBy(match => match.Variation.DigiKeyProductNumber, StringComparer.OrdinalIgnoreCase)
                .First();

            return (bestMatch.Product, bestMatch.Variation);
        }

        private static List<DigiKeyProduct> CollectSearchProducts(DigiKeyKeywordSearchResponse? searchResponse)
        {
            var products = new List<DigiKeyProduct>();
            var seenManufacturerPartNumbers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AddProducts(IEnumerable<DigiKeyProduct>? candidates)
            {
                if (candidates == null)
                {
                    return;
                }

                foreach (var product in candidates)
                {
                    var key = product.ManufacturerProductNumber?.Trim();
                    if (string.IsNullOrWhiteSpace(key))
                    {
                        products.Add(product);
                        continue;
                    }

                    if (seenManufacturerPartNumbers.Add(key))
                    {
                        products.Add(product);
                    }
                }
            }

            AddProducts(searchResponse?.ExactMatches);
            AddProducts(searchResponse?.Products);
            return products;
        }

        private static bool IsExcludedReelOnlyProduct(DigiKeyProduct product)
        {
            var productPackagingValues = GetProductPackagingValues(product).ToList();
            if (productPackagingValues.Count == 0)
            {
                return false;
            }

            var hasCutTapeIndicator = productPackagingValues.Any(ContainsCutTapeIndicator);
            if (hasCutTapeIndicator)
            {
                return false;
            }

            return productPackagingValues.All(IsTapeAndReelPackaging);
        }

        private static DigiKeyProductVariation? FindCutTapeVariation(
            DigiKeyProduct product,
            string searchedPartNumber)
        {
            var variations = product.ProductVariations;
            if (variations == null || variations.Count == 0)
            {
                return null;
            }

            var cutTapeVariations = variations
                .Where(variation => IsCutTapePackaging(variation) && !IsTapeAndReelPackaging(variation))
                .ToList();

            if (cutTapeVariations.Count == 0)
            {
                return null;
            }

            var exactPartNumberMatch = cutTapeVariations.FirstOrDefault(variation =>
                string.Equals(variation.DigiKeyProductNumber, searchedPartNumber, StringComparison.OrdinalIgnoreCase));
            if (exactPartNumberMatch != null)
            {
                return exactPartNumberMatch;
            }

            return cutTapeVariations
                .OrderByDescending(variation => IsCutTapeDigiKeyPartNumber(variation.DigiKeyProductNumber))
                .ThenBy(variation => variation.DigiKeyProductNumber, StringComparer.OrdinalIgnoreCase)
                .First();
        }

        private static int ScoreCutTapeMatch(
            DigiKeyProduct product,
            DigiKeyProductVariation variation,
            string searchedPartNumber,
            string searchedManufacturer)
        {
            var score = 0;

            if (GetProductPackagingValues(product).Any(ContainsCutTapeIndicator))
            {
                score += 100;
            }

            if (IsCutTapePackaging(variation))
            {
                score += 50;
            }

            if (!string.IsNullOrWhiteSpace(searchedManufacturer))
            {
                var digiKeyManufacturer = product.Manufacturer?.Name ?? string.Empty;
                if (digiKeyManufacturer.Equals(searchedManufacturer, StringComparison.OrdinalIgnoreCase))
                {
                    score += 80;
                }
                else if (digiKeyManufacturer.Contains(searchedManufacturer, StringComparison.OrdinalIgnoreCase)
                    || searchedManufacturer.Contains(digiKeyManufacturer, StringComparison.OrdinalIgnoreCase))
                {
                    score += 40;
                }
            }

            if (string.Equals(variation.DigiKeyProductNumber, searchedPartNumber, StringComparison.OrdinalIgnoreCase))
            {
                score += 40;
            }

            if (string.Equals(product.ManufacturerProductNumber, searchedPartNumber, StringComparison.OrdinalIgnoreCase))
            {
                score += 30;
            }

            if (IsCutTapeDigiKeyPartNumber(variation.DigiKeyProductNumber))
            {
                score += 20;
            }

            if (IsCutTapeDigiKeyPartNumber(searchedPartNumber)
                && variation.DigiKeyProductNumber?.Contains("CT", StringComparison.OrdinalIgnoreCase) == true)
            {
                score += 10;
            }

            return score;
        }

        private static IEnumerable<string> GetProductPackagingValues(DigiKeyProduct product)
        {
            if (!string.IsNullOrWhiteSpace(product.Packaging?.Value))
            {
                yield return product.Packaging.Value;
            }

            if (!string.IsNullOrWhiteSpace(product.PackagingType?.Value))
            {
                yield return product.PackagingType.Value;
            }

            if (!string.IsNullOrWhiteSpace(product.PackageType?.Name))
            {
                yield return product.PackageType.Name;
            }
        }

        private static string GetVariationPackagingName(DigiKeyProductVariation variation)
        {
            return variation.PackageType?.Name
                ?? variation.Packaging
                ?? variation.PackagingType
                ?? string.Empty;
        }

        private static bool IsCutTapePackaging(DigiKeyProductVariation variation)
        {
            if (variation.PackageType?.Id == CutTapePackageTypeId)
            {
                return true;
            }

            return ContainsCutTapeIndicator(GetVariationPackagingName(variation));
        }

        private static bool IsTapeAndReelPackaging(DigiKeyProductVariation variation)
        {
            if (variation.PackageType?.Id == TapeAndReelPackageTypeId)
            {
                return true;
            }

            return IsTapeAndReelPackaging(GetVariationPackagingName(variation));
        }

        private static bool ContainsCutTapeIndicator(string? packagingName)
        {
            if (string.IsNullOrWhiteSpace(packagingName) || IsTapeAndReelPackaging(packagingName))
            {
                return false;
            }

            return packagingName.Equals("Cut Tape", StringComparison.OrdinalIgnoreCase)
                || packagingName.Equals("CT", StringComparison.OrdinalIgnoreCase)
                || packagingName.Contains("Cut Tape (CT)", StringComparison.OrdinalIgnoreCase)
                || packagingName.Contains("Cut Tape", StringComparison.OrdinalIgnoreCase)
                || packagingName.Contains("(CT)", StringComparison.OrdinalIgnoreCase)
                || packagingName.Contains("컷 테이프", StringComparison.OrdinalIgnoreCase)
                || packagingName.Contains("컷테이프", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTapeAndReelPackaging(string? packagingName)
        {
            if (string.IsNullOrWhiteSpace(packagingName))
            {
                return false;
            }

            if (packagingName.Contains("Tape & Reel", StringComparison.OrdinalIgnoreCase)
                || packagingName.Contains("Tape and Reel", StringComparison.OrdinalIgnoreCase)
                || packagingName.Contains("Tape &amp; Reel", StringComparison.OrdinalIgnoreCase)
                || packagingName.Contains("테이프 & 릴", StringComparison.OrdinalIgnoreCase)
                || packagingName.Contains("테이프 릴", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (packagingName.Equals("TR", StringComparison.OrdinalIgnoreCase)
                || packagingName.Equals("Reel", StringComparison.OrdinalIgnoreCase)
                || packagingName.Equals("릴", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (packagingName.Contains("Reel", StringComparison.OrdinalIgnoreCase)
                && !packagingName.Contains("Cut Tape", StringComparison.OrdinalIgnoreCase)
                && !packagingName.Contains("Digi-Reel", StringComparison.OrdinalIgnoreCase)
                && !packagingName.Contains("컷", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private static bool IsCutTapeDigiKeyPartNumber(string? digiKeyPartNumber)
        {
            if (string.IsNullOrWhiteSpace(digiKeyPartNumber))
            {
                return false;
            }

            return digiKeyPartNumber.Contains("CT-ND", StringComparison.OrdinalIgnoreCase)
                || digiKeyPartNumber.EndsWith("CT", StringComparison.OrdinalIgnoreCase);
        }

        private static DigiKeyProductSummary MapProductSummary(
            string searchedPartNumber,
            string searchedManufacturer,
            DigiKeyProduct product,
            DigiKeyProductVariation variation)
        {
            var sortedPricing = SortPricingByBreakQuantity(variation.StandardPricing);
            var quantityAvailable = variation.QuantityAvailableForPackageType > 0
                ? variation.QuantityAvailableForPackageType
                : product.QuantityAvailable;

            return new DigiKeyProductSummary
            {
                SearchedPartNumber = searchedPartNumber,
                SearchedManufacturer = string.IsNullOrWhiteSpace(searchedManufacturer) ? "-" : searchedManufacturer,
                DigiKeyPartNumber = variation.DigiKeyProductNumber ?? "-",
                DigiKeyManufacturer = product.Manufacturer?.Name ?? "-",
                PackagingType = "Cut Tape (CT)",
                QuantityAvailable = quantityAvailable,
                PriceTiers = ParsePriceTiers(sortedPricing)
            };
        }

        private static List<DigiKeyPriceTier> ParsePriceTiers(IReadOnlyList<DigiKeyPriceBreak> sortedPricing)
        {
            return sortedPricing
                .Select(priceBreak => new DigiKeyPriceTier
                {
                    BreakQuantity = priceBreak.BreakQuantity,
                    FormattedUnitPrice = FormatKrwPrice(priceBreak.UnitPrice)
                })
                .ToList();
        }

        private static List<DigiKeyPriceBreak> SortPricingByBreakQuantity(
            IReadOnlyList<DigiKeyPriceBreak>? pricing)
        {
            return pricing?.OrderBy(price => price.BreakQuantity).ToList() ?? new List<DigiKeyPriceBreak>();
        }

        private static string FormatKrwPrice(decimal unitPrice)
        {
            return "₩" + unitPrice.ToString("#,##0", CultureInfo.GetCultureInfo("ko-KR"));
        }

        private static string ExtractApiError(string body)
        {
            if (string.IsNullOrWhiteSpace(body))
            {
                return "응답 본문이 비어 있습니다.";
            }

            try
            {
                var problem = JsonSerializer.Deserialize<DigiKeyProblemDetails>(body, JsonOptions);
                if (!string.IsNullOrWhiteSpace(problem?.Detail))
                {
                    return problem.Detail;
                }

                if (!string.IsNullOrWhiteSpace(problem?.Title))
                {
                    return problem.Title;
                }
            }
            catch (JsonException)
            {
            }

            return body.Length > 300 ? body[..300] + "..." : body;
        }
    }

    internal sealed class DigiKeyProductSummary
    {
        public string SearchedPartNumber { get; init; } = string.Empty;
        public string SearchedManufacturer { get; init; } = string.Empty;
        public string DigiKeyPartNumber { get; init; } = string.Empty;
        public string DigiKeyManufacturer { get; init; } = string.Empty;
        public string PackagingType { get; init; } = string.Empty;
        public int QuantityAvailable { get; init; }
        public IReadOnlyList<DigiKeyPriceTier> PriceTiers { get; init; } = Array.Empty<DigiKeyPriceTier>();
    }

    internal sealed class DigiKeyPriceTier
    {
        public int BreakQuantity { get; init; }
        public string FormattedUnitPrice { get; init; } = "N/A";
    }

    internal sealed class DigiKeyKeywordSearchRequest
    {
        [JsonPropertyName("Keywords")]
        public string Keywords { get; set; } = string.Empty;

        [JsonPropertyName("Limit")]
        public int Limit { get; set; }

        [JsonPropertyName("Offset")]
        public int Offset { get; set; }
    }

    internal sealed class DigiKeyTokenResponse
    {
        [JsonPropertyName("access_token")]
        public string? AccessToken { get; set; }
    }

    internal sealed class DigiKeyKeywordSearchResponse
    {
        [JsonPropertyName("Products")]
        public List<DigiKeyProduct>? Products { get; set; }

        [JsonPropertyName("ExactMatches")]
        public List<DigiKeyProduct>? ExactMatches { get; set; }
    }

    internal sealed class DigiKeyProduct
    {
        [JsonPropertyName("Description")]
        public DigiKeyDescription? Description { get; set; }

        [JsonPropertyName("Manufacturer")]
        public DigiKeyManufacturer? Manufacturer { get; set; }

        [JsonPropertyName("ManufacturerProductNumber")]
        public string? ManufacturerProductNumber { get; set; }

        [JsonPropertyName("QuantityAvailable")]
        public int QuantityAvailable { get; set; }

        [JsonPropertyName("Packaging")]
        public DigiKeyPackagingValue? Packaging { get; set; }

        [JsonPropertyName("PackagingType")]
        public DigiKeyPackagingValue? PackagingType { get; set; }

        [JsonPropertyName("PackageType")]
        public DigiKeyPackageType? PackageType { get; set; }

        [JsonPropertyName("ProductVariations")]
        public List<DigiKeyProductVariation>? ProductVariations { get; set; }
    }

    internal sealed class DigiKeyDescription
    {
        [JsonPropertyName("ProductDescription")]
        public string? ProductDescription { get; set; }

        [JsonPropertyName("DetailedDescription")]
        public string? DetailedDescription { get; set; }
    }

    internal sealed class DigiKeyManufacturer
    {
        [JsonPropertyName("Name")]
        public string? Name { get; set; }
    }

    internal sealed class DigiKeyPackagingValue
    {
        [JsonPropertyName("Id")]
        public int Id { get; set; }

        [JsonPropertyName("Value")]
        public string? Value { get; set; }
    }

    internal sealed class DigiKeyProductVariation
    {
        [JsonPropertyName("DigiKeyProductNumber")]
        public string? DigiKeyProductNumber { get; set; }

        [JsonPropertyName("PackageType")]
        public DigiKeyPackageType? PackageType { get; set; }

        [JsonPropertyName("Packaging")]
        public string? Packaging { get; set; }

        [JsonPropertyName("PackagingType")]
        public string? PackagingType { get; set; }

        [JsonPropertyName("MinimumOrderQuantity")]
        public int MinimumOrderQuantity { get; set; }

        [JsonPropertyName("QuantityAvailableForPackageType")]
        public int QuantityAvailableForPackageType { get; set; }

        [JsonPropertyName("StandardPricing")]
        public List<DigiKeyPriceBreak>? StandardPricing { get; set; }
    }

    internal sealed class DigiKeyPackageType
    {
        [JsonPropertyName("Id")]
        public int Id { get; set; }

        [JsonPropertyName("Name")]
        public string? Name { get; set; }
    }

    internal sealed class DigiKeyPriceBreak
    {
        [JsonPropertyName("BreakQuantity")]
        public int BreakQuantity { get; set; }

        [JsonPropertyName("UnitPrice")]
        public decimal UnitPrice { get; set; }
    }

    internal sealed class DigiKeyProblemDetails
    {
        [JsonPropertyName("title")]
        public string? Title { get; set; }

        [JsonPropertyName("detail")]
        public string? Detail { get; set; }
    }
}
