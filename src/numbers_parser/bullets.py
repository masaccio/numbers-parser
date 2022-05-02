from roman import toRoman
from numbers_parser.generated.TSWPArchives_pb2 import ListStyleArchive


BULLET_PREFIXES = {
    ListStyleArchive.kNumericDecimal: "",
    ListStyleArchive.kNumericDoubleParen: "(",
    ListStyleArchive.kNumericRightParen: "",
    ListStyleArchive.kRomanUpperDecimal: "",
    ListStyleArchive.kRomanUpperDoubleParen: "(",
    ListStyleArchive.kRomanUpperRightParen: "",
    ListStyleArchive.kRomanLowerDecimal: "",
    ListStyleArchive.kRomanLowerDoubleParen: "(",
    ListStyleArchive.kRomanLowerRightParen: "",
    ListStyleArchive.kAlphaUpperDecimal: "",
    ListStyleArchive.kAlphaUpperDoubleParen: "(",
    ListStyleArchive.kAlphaUpperRightParen: "",
    ListStyleArchive.kAlphaLowerDecimal: "",
    ListStyleArchive.kAlphaLowerDoubleParen: "(",
    ListStyleArchive.kAlphaLowerRightParen: "",
}

BULLET_CONVERTION = {
    ListStyleArchive.kNumericDecimal: lambda x: str(x + 1),
    ListStyleArchive.kNumericDoubleParen: lambda x: str(x + 1),
    ListStyleArchive.kNumericRightParen: lambda x: str(x + 1),
    ListStyleArchive.kRomanUpperDecimal: lambda x: toRoman(x + 1),
    ListStyleArchive.kRomanUpperDoubleParen: lambda x: toRoman(x + 1),
    ListStyleArchive.kRomanUpperRightParen: lambda x: toRoman(x + 1),
    ListStyleArchive.kRomanLowerDecimal: lambda x: toRoman(x + 1).lower(),
    ListStyleArchive.kRomanLowerDoubleParen: lambda x: toRoman(x + 1).lower(),
    ListStyleArchive.kRomanLowerRightParen: lambda x: toRoman(x + 1).lower(),
    ListStyleArchive.kAlphaUpperDecimal: lambda x: chr(x + 65),
    ListStyleArchive.kAlphaUpperDoubleParen: lambda x: chr(x + 65),
    ListStyleArchive.kAlphaUpperRightParen: lambda x: chr(x + 65),
    ListStyleArchive.kAlphaLowerDecimal: lambda x: chr(x + 97),
    ListStyleArchive.kAlphaLowerDoubleParen: lambda x: chr(x + 97),
    ListStyleArchive.kAlphaLowerRightParen: lambda x: chr(x + 97),
}

BULLET_SUFFIXES = {
    ListStyleArchive.kNumericDecimal: ".",
    ListStyleArchive.kNumericDoubleParen: ")",
    ListStyleArchive.kNumericRightParen: ")",
    ListStyleArchive.kRomanUpperDecimal: ".",
    ListStyleArchive.kRomanUpperDoubleParen: ")",
    ListStyleArchive.kRomanUpperRightParen: ")",
    ListStyleArchive.kRomanLowerDecimal: ".",
    ListStyleArchive.kRomanLowerDoubleParen: ")",
    ListStyleArchive.kRomanLowerRightParen: ")",
    ListStyleArchive.kAlphaUpperDecimal: ".",
    ListStyleArchive.kAlphaUpperDoubleParen: ")",
    ListStyleArchive.kAlphaUpperRightParen: ")",
    ListStyleArchive.kAlphaLowerDecimal: ".",
    ListStyleArchive.kAlphaLowerDoubleParen: ")",
    ListStyleArchive.kAlphaLowerRightParen: ")",
}
