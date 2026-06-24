FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY _server.py _extract_openpyxl.py HALO_Dashboard.html ./
COPY _data_block.js* ./
RUN mkdir -p "Ops files"
ENV HALO_CLOUD=1
ENV PORT=8765
EXPOSE 8765
CMD ["python", "_server.py"]
