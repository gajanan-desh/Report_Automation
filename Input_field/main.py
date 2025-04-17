from flask import Flask, request, send_file, render_template
from docxtpl import DocxTemplate
import os
from datetime import datetime
import io

app = Flask(__name__)

@app.route('/')
def index():
 
    return render_template('form.html')

@app.route('/generate', methods=['POST'])
def generate_document():
  
    company_name = request.form.get('companyName', 'Company Name')
    network_type = request.form.get('networkType', 'External Network')
    assessment_date = request.form.get('assessmentDate', datetime.now().strftime('%d/%m/%Y'))
    findings_count = request.form.get('findingsCount', '5')
    
   
    template_path = os.path.join(app.root_path, 'templates', 'report_template.docx')
    doc = DocxTemplate(template_path)
    
   
    context = {
        'companyName': company_name,
        'networkType': network_type,
        'assessmentDate': assessment_date,
        'findingsCount': findings_count,
        
    }
    
    
    doc.render(context)
    
   
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    
   
    filename = f"{company_name}_Assessment_Report.docx"
    return send_file(
        output, 
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

if __name__ == '__main__':
    app.run(debug=True)