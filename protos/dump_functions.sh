TSCE=${1:-/Applications/Numbers.app/Contents/Frameworks/TSCalculationEngine.framework/Versions/A/TSCalculationEngine}

export LC_CTYPE=C
echo 'FUNCTION_MAP = {'
arch -x86_64 objdump -macho -objc-meta-data -disassemble "$TSCE" | awk '/movl/ && /24\(%rsp\)/ { key = $(NF-1); }
/movl/ && / \(%rsp\)/ { op = $(NF-1) == "$1," ? 1 : 0; }
/leaq/ && /cfstring ref/ { name = $NF; }
/TSCEFunction(Argument)?Spec (specWithFunctionName|argSpecWithType|setObject)/ && !op {
	print key, name;
	name = key = ""; op = 0;
}' | sed 's/\$//g; s/,/:/g; s/@//g; s/$/,/g; s/^/    /'

echo '}'
