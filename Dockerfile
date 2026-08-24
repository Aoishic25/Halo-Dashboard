FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY _server.py _extract_openpyxl.py index.html ./
COPY compromise.min.js ./
COPY _data_block.js* ./
COPY Theme.json ./
RUN mkdir -p "Ops files"
ENV HALO_CLOUD=1
ENV PORT=8765
EXPOSE 8765
CMD ["python", "_server.py"]
