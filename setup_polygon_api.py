"""
Deploy coverage_polygons API to the Laravel VPS.

Usage:
    export VPS_HOST=<ip>  VPS_SSH_PASS=<pass>  VPS_MYSQL_PASS=<pass>
    python setup_polygon_api.py
"""
import paramiko
import posixpath

import os
host      = os.environ.get("VPS_HOST",       "YOUR_VPS_IP")
user      = os.environ.get("VPS_USER",       "root")
ssh_pass  = os.environ.get("VPS_SSH_PASS",   "")
mysql_pass= os.environ.get("VPS_MYSQL_PASS", "")
base_path = os.environ.get("VPS_BASE_PATH",  "/var/www/epi-api")

client = paramiko.SSHClient()
client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
client.connect(host, username=user, password=ssh_pass, timeout=30)

def run(cmd):
    stdin, stdout, stderr = client.exec_command(cmd)
    out = stdout.read().decode()
    err = stderr.read().decode()
    if out.strip(): print("  OUT:", out.strip())
    if err.strip(): print("  ERR:", err.strip())
    return out, err

def write_file(path, content):
    sftp = client.open_sftp()
    dir_path = posixpath.dirname(path)
    run(f"mkdir -p {dir_path}")
    with sftp.file(path, 'w') as f:
        f.write(content)
    sftp.close()
    print(f"  Written: {path}")

print("=== Connected to VPS — deploying Polygon API ===\n")

# ── 1. Migration ──────────────────────────────────────────
migration = """<?php
use Illuminate\\Database\\Migrations\\Migration;
use Illuminate\\Database\\Schema\\Blueprint;
use Illuminate\\Support\\Facades\\Schema;

return new class extends Migration {
    public function up(): void {
        Schema::create('coverage_polygons', function (Blueprint $table) {
            $table->id();
            $table->string('name')->default('Untitled Polygon');
            $table->string('created_by')->nullable();
            $table->string('updated_by')->nullable();
            $table->text('notes')->nullable();
            $table->string('governorate')->nullable();
            $table->string('color', 20)->default('#7E57C2');
            $table->string('line_color', 20)->nullable();
            $table->unsignedTinyInteger('line_weight')->default(2);
            $table->float('opacity')->default(0.35);
            $table->float('fill_opacity')->default(0.18);
            $table->json('points');
            $table->string('source', 50)->default('api');
            $table->timestamps();
        });
    }
    public function down(): void {
        Schema::dropIfExists('coverage_polygons');
    }
};
"""

# ── 2. Model ──────────────────────────────────────────────
model = """<?php
namespace App\\Models;
use Illuminate\\Database\\Eloquent\\Model;

class CoveragePolygon extends Model {
    protected $fillable = [
        'name','created_by','updated_by','notes','governorate',
        'color','line_color','line_weight','opacity','fill_opacity','points','source'
    ];
    protected $casts = ['points' => 'array'];
}
"""

# ── 3. Controller ─────────────────────────────────────────
controller = """<?php
namespace App\\Http\\Controllers;

use App\\Models\\CoveragePolygon;
use Illuminate\\Http\\Request;

class CoveragePolygonController extends Controller
{
    public function index()
    {
        return response()->json(
            CoveragePolygon::orderByDesc('updated_at')->orderByDesc('id')->get()
        );
    }

    public function store(Request $request)
    {
        $data = $request->validate([
            'name'         => 'nullable|string|max:255',
            'created_by'   => 'nullable|string|max:100',
            'updated_by'   => 'nullable|string|max:100',
            'notes'        => 'nullable|string',
            'governorate'  => 'nullable|string|max:100',
            'color'        => 'nullable|string|max:20',
            'line_color'   => 'nullable|string|max:20',
            'line_weight'  => 'nullable|integer',
            'opacity'      => 'nullable|numeric',
            'fill_opacity' => 'nullable|numeric',
            'points'       => 'required|array|min:3',
            'source'       => 'nullable|string|max:50',
        ]);
        $polygon = CoveragePolygon::create($data);
        return response()->json($polygon, 201);
    }

    public function destroy($id)
    {
        $polygon = CoveragePolygon::findOrFail($id);
        $polygon->delete();
        return response()->json(['deleted' => true]);
    }
}
"""

# ── 4. API Routes snippet ─────────────────────────────────
routes_snippet = """
// Coverage Polygons — multi-user drawing
Route::get('/polygons',       [App\\Http\\Controllers\\CoveragePolygonController::class, 'index']);
Route::post('/polygons',      [App\\Http\\Controllers\\CoveragePolygonController::class, 'store']);
Route::delete('/polygons/{id}', [App\\Http\\Controllers\\CoveragePolygonController::class, 'destroy']);
"""

# ── 5. CORS config ────────────────────────────────────────
cors_config = """<?php
return [
    'paths' => ['api/*'],
    'allowed_methods' => ['*'],
    'allowed_origins' => ['*'],
    'allowed_origins_patterns' => [],
    'allowed_headers' => ['*'],
    'exposed_headers' => [],
    'max_age' => 0,
    'supports_credentials' => false,
];
"""

# Write files
print("Writing migration...")
write_file(f"{base_path}/database/migrations/2025_01_01_000001_create_coverage_polygons_table.php", migration)

print("Writing model...")
write_file(f"{base_path}/app/Models/CoveragePolygon.php", model)

print("Writing controller...")
write_file(f"{base_path}/app/Http/Controllers/CoveragePolygonController.php", controller)

print("Writing CORS config...")
write_file(f"{base_path}/config/cors.php", cors_config)

print("Appending routes...")
run(f"grep -q 'Coverage Polygons' {base_path}/routes/api.php || echo '{routes_snippet}' >> {base_path}/routes/api.php")

print("\nRunning migration...")
out, err = run(f"cd {base_path} && php artisan migrate --force 2>&1")
print(out or err)

print("Clearing caches...")
run(f"cd {base_path} && php artisan config:clear && php artisan route:clear && php artisan cache:clear")

print("\n=== Done! Test with: curl http://144.172.102.6/api/polygons ===")
client.close()
