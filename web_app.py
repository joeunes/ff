import os
import time
import shutil
import tempfile
from flask import Flask, request, render_template, send_file, jsonify
from flask_cors import CORS
from hwp_autowriter import run_bim_to_excel

app = Flask(__name__)
CORS(app)

# Configure upload and output directories
if os.name == 'nt':
    UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
    OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs')
else:
    UPLOAD_FOLDER = '/tmp/uploads'
    OUTPUT_FOLDER = '/tmp/outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    file_a = request.files.get('file_a')
    file_s = request.files.get('file_s')
    
    if not file_a and not file_s:
        return jsonify({'error': '적어도 하나의 IFC 파일을 업로드해야 합니다.'}), 400
        
    temp_dir = tempfile.mkdtemp(dir=UPLOAD_FOLDER)
    
    path_a = ""
    path_s = ""
    
    try:
        # Save BA.ifc if uploaded
        if file_a and file_a.filename:
            path_a = os.path.join(temp_dir, 'BA.ifc')
            file_a.save(path_a)
            
        # Save BS.ifc if uploaded
        if file_s and file_s.filename:
            path_s = os.path.join(temp_dir, 'BS.ifc')
            file_s.save(path_s)
            
        # Define output excel file path
        output_filename = f"BIM_수량_및_공간_산출서_{int(time.time())}.xlsx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        # Run calculation engine
        print(f"[Web Webapp] Starting quantity takeoff calculation...")
        print(f"  Architectural path: {path_a}")
        print(f"  Structural path: {path_s}")
        print(f"  Output path: {output_path}")
        
        summary = run_bim_to_excel(path_a, path_s, output_path)
        actual_path = summary.get("saved_excel_path", output_path)
        output_filename = os.path.basename(actual_path)
        
        if not os.path.exists(actual_path):
            return jsonify({'error': '엑셀 파일을 생성하지 못했습니다.'}), 500
            
        # Check if geometry JSON files were created
        arch_geom_filename = output_filename.replace(".xlsx", "_arch.json")
        struct_geom_filename = output_filename.replace(".xlsx", "_struct.json")
        
        arch_geom_url = f"/geom/{arch_geom_filename}" if os.path.exists(os.path.join(OUTPUT_FOLDER, arch_geom_filename)) else ""
        struct_geom_url = f"/geom/{struct_geom_filename}" if os.path.exists(os.path.join(OUTPUT_FOLDER, struct_geom_filename)) else ""
            
        return jsonify({
            'success': True,
            'download_url': f'/download/{output_filename}',
            'filename': output_filename,
            'arch_geom_url': arch_geom_url,
            'struct_geom_url': struct_geom_url,
            'readouts': {
                'arch_wall_area': f"{summary['arch_wall_area']:,.1f} m²" if summary.get('arch_wall_area') else "—",
                'arch_space_area': f"{summary['arch_space_area']:,.1f} m²" if summary.get('arch_space_area') else "—",
                'struct_element_qty': f"{summary['struct_element_qty']:,} EA" if summary.get('struct_element_qty') else "—"
            }
        })
        
    except Exception as e:
        print(f"[Web Webapp] Error occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'파일 처리 중 오류가 발생했습니다: {str(e)}'}), 500
        
    finally:
        # Cleanup uploaded raw files
        try:
            shutil.rmtree(temp_dir)
        except Exception:
            pass

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(file_path):
        return "파일을 찾을 수 없습니다.", 404
        
    # Return the file and arrange for it to be deleted after sending
    response = send_file(file_path, as_attachment=True, download_name="BIM_수량_및_공간_산출서.xlsx")
    
    # We can delete the file after some time or using after_this_request
    @response.call_on_close
    def cleanup():
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"[Web Webapp] Cleaned up output file: {file_path}")
        except Exception as e:
            print(f"[Web Webapp] Cleanup error: {str(e)}")
            
    return response

@app.route('/geom/<filename>')
def get_geom(filename):
    file_path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(file_path):
        return "파일을 찾을 수 없습니다.", 404
        
    response = send_file(file_path, mimetype='application/json')
    
    @response.call_on_close
    def cleanup():
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"[Web Webapp] Cleaned up geom file: {file_path}")
        except Exception as e:
            print(f"[Web Webapp] Cleanup error: {str(e)}")
            
    return response

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
