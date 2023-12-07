from transformers import GPT2LMHeadModel, GPT2Tokenizer

log_file = "C:/Users/kevin/OneDrive/Escritorio/BOT - BASE/prueba.log"  # Nombre del archivo de registro

# Leer el contenido del archivo de registro
with open(log_file, "r") as file:
    text = file.read()

# Paso 2: Cargar el modelo y el tokenizador
model_name = 'gpt2'  # Puedes elegir otros modelos según tus necesidades
model = GPT2LMHeadModel.from_pretrained(model_name)
tokenizer = GPT2Tokenizer.from_pretrained(model_name)

# Verificar si el tokenizador tiene un token de relleno definido
if tokenizer.pad_token is None:
    # Establecer un token de relleno genérico
    tokenizer.add_special_tokens({'pad_token': '[PAD]'})
    model.resize_token_embeddings(len(tokenizer))

# Paso 3: Dividir el texto en segmentos más pequeños
max_length = model.config.n_positions - 2  # Restar 2 para dejar espacio para los tokens especiales
chunks = [text[i:i + max_length] for i in range(0, len(text), max_length)]

# Paso 4: Generar la interpretación para cada segmento
interpretations = []
for chunk in chunks:
    input_ids = tokenizer.encode(chunk, return_tensors='pt', truncation=True, max_length=max_length)
    attention_mask = input_ids.ne(tokenizer.pad_token_id) if tokenizer.pad_token_id is not None else None
    output = model.generate(input_ids, attention_mask=attention_mask, max_length=512, num_return_sequences=1)
    interpretation = tokenizer.decode(output[0], skip_special_tokens=True)
    interpretations.append(interpretation)

# Paso 5: Concatenar las interpretaciones generadas
full_interpretation = " ".join(interpretations)

# Paso 6: Presentar la interpretación completa
print("Interpretación del texto:")
print(full_interpretation)
