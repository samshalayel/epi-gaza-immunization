import paramiko
import sys

host = "144.172.102.6"
user = "root"
password = "21vU9xtxSVyFt3"
base_path = "/var/www/epi-api"

client = paramiko.SSHClient()
client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
client.connect(host, username=user, password=password, timeout=30)

def run(cmd):
    stdin, stdout, stderr = client.exec_command(cmd)
    out = stdout.read().decode()
    err = stderr.read().decode()
    return out, err

def write_file(path, content):
    """Write a file on the remote server using SFTP, creating dirs if needed."""
    sftp = client.open_sftp()
    # Ensure parent directory exists
    import posixpath
    dir_path = posixpath.dirname(path)
    # mkdir -p equivalent via SSH
    run(f"mkdir -p {dir_path}")
    with sftp.file(path, 'w') as f:
        f.write(content)
    sftp.close()
    print(f"  Written: {path}")

print("Connected to VPS. Creating Laravel API files...\n")

# 1. app/Models/Facility.php
facility_model = """<?php
namespace App\\Models;
use Illuminate\\Database\\Eloquent\\Model;
class Facility extends Model {
    protected $fillable = ['name','governorate','type','provider','is_functional'];
    protected $casts = ['is_functional' => 'boolean'];
    public function data() { return $this->hasMany(FacilityData::class); }
}
"""

# 2. app/Models/FacilityData.php
facility_data_model = """<?php
namespace App\\Models;
use Illuminate\\Database\\Eloquent\\Model;
class FacilityData extends Model {
    protected $fillable = ['facility_id','year','catchment_pop','si_percent','monthly_data'];
    protected $casts = ['monthly_data' => 'array'];
    public function facility() { return $this->belongsTo(Facility::class); }
}
"""

# 3. app/Http/Controllers/FacilityController.php
facility_controller = """<?php
namespace App\\Http\\Controllers;

use App\\Models\\Facility;
use Illuminate\\Http\\Request;

class FacilityController extends Controller
{
    public function index()
    {
        return response()->json(Facility::orderBy('name')->get());
    }

    public function store(Request $request)
    {
        $validated = $request->validate(['name' => 'required|string']);
        $facility = Facility::create($request->all());
        return response()->json($facility, 201);
    }

    public function update(Request $request, $id)
    {
        $facility = Facility::findOrFail($id);
        $facility->update($request->all());
        return response()->json($facility);
    }

    public function destroy($id)
    {
        $facility = Facility::findOrFail($id);
        $facility->delete();
        return response()->json(null, 204);
    }

    public function import(Request $request)
    {
        $facilities = $request->json()->all();
        $count = 0;
        foreach ($facilities as $item) {
            Facility::updateOrCreate(
                ['name' => $item['name']],
                [
                    'governorate' => $item['governorate'] ?? null,
                    'type'        => $item['type'] ?? null,
                    'provider'    => $item['provider'] ?? null,
                ]
            );
            $count++;
        }
        return response()->json(['imported' => $count]);
    }
}
"""

# 4. app/Http/Controllers/FacilityDataController.php
facility_data_controller = """<?php
namespace App\\Http\\Controllers;

use App\\Models\\FacilityData;
use Illuminate\\Http\\Request;

class FacilityDataController extends Controller
{
    public function show($facilityId, $year)
    {
        $data = FacilityData::where('facility_id', $facilityId)
                            ->where('year', $year)
                            ->first();
        if (!$data) {
            return response()->json((object)[]);
        }
        return response()->json($data);
    }

    public function store(Request $request, $facilityId, $year)
    {
        $data = FacilityData::updateOrCreate(
            ['facility_id' => $facilityId, 'year' => $year],
            [
                'catchment_pop' => $request->input('catchment_pop'),
                'si_percent'    => $request->input('si_percent'),
                'monthly_data'  => $request->input('monthly_data'),
            ]
        );
        return response()->json($data);
    }

    public function index($facilityId)
    {
        $data = FacilityData::where('facility_id', $facilityId)
                            ->orderBy('year')
                            ->get();
        return response()->json($data);
    }
}
"""

# 5. routes/api.php
api_routes = """<?php
use App\\Http\\Controllers\\FacilityController;
use App\\Http\\Controllers\\FacilityDataController;
use Illuminate\\Support\\Facades\\Route;

Route::get('/facilities', [FacilityController::class, 'index']);
Route::post('/facilities/import', [FacilityController::class, 'import']);
Route::post('/facilities', [FacilityController::class, 'store']);
Route::put('/facilities/{id}', [FacilityController::class, 'update']);
Route::delete('/facilities/{id}', [FacilityController::class, 'destroy']);

Route::get('/data/{facilityId}/{year}', [FacilityDataController::class, 'show']);
Route::post('/data/{facilityId}/{year}', [FacilityDataController::class, 'store']);
Route::get('/data/{facilityId}', [FacilityDataController::class, 'index']);
"""

# 6. CORS Middleware
cors_middleware = """<?php
namespace App\\Http\\Middleware;

use Closure;
use Illuminate\\Http\\Request;

class Cors
{
    public function handle(Request $request, Closure $next)
    {
        if ($request->isMethod('OPTIONS')) {
            return response('', 200)
                ->header('Access-Control-Allow-Origin', '*')
                ->header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS')
                ->header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
        }

        $response = $next($request);
        $response->headers->set('Access-Control-Allow-Origin', '*');
        $response->headers->set('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
        $response->headers->set('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
        return $response;
    }
}
"""

# 7. bootstrap/app.php (Laravel 11 style with CORS + api routing)
bootstrap_app = """<?php

use Illuminate\\Foundation\\Application;
use Illuminate\\Foundation\\Configuration\\Exceptions;
use Illuminate\\Foundation\\Configuration\\Middleware;
use App\\Http\\Middleware\\Cors;

return Application::configure(basePath: dirname(__DIR__))
    ->withRouting(
        web: __DIR__.'/../routes/web.php',
        api: __DIR__.'/../routes/api.php',
        commands: __DIR__.'/../routes/console.php',
        health: '/up',
    )
    ->withMiddleware(function (Middleware $middleware): void {
        $middleware->prepend(Cors::class);
    })
    ->withExceptions(function (Exceptions $exceptions): void {
        //
    })->create();
"""

# Write all files
print("Writing model files...")
write_file(f"{base_path}/app/Models/Facility.php", facility_model)
write_file(f"{base_path}/app/Models/FacilityData.php", facility_data_model)

print("Writing controller files...")
write_file(f"{base_path}/app/Http/Controllers/FacilityController.php", facility_controller)
write_file(f"{base_path}/app/Http/Controllers/FacilityDataController.php", facility_data_controller)

print("Writing routes/api.php...")
write_file(f"{base_path}/routes/api.php", api_routes)

print("Writing CORS middleware...")
write_file(f"{base_path}/app/Http/Middleware/Cors.php", cors_middleware)

print("Writing bootstrap/app.php (Laravel 11 with api routing + CORS)...")
write_file(f"{base_path}/bootstrap/app.php", bootstrap_app)

# Clear all caches
print("\nClearing Laravel caches...")
out, err = run(f"cd {base_path} && php artisan config:clear 2>&1 && php artisan route:clear 2>&1 && php artisan cache:clear 2>&1")
print(out)
if err:
    print("STDERR:", err)

# Verify routes
print("\n=== ROUTE LIST: facilities and data routes ===")
out, err = run(f"cd {base_path} && php artisan route:list 2>&1 | grep -i 'facilit\\|/data'")
print(out)
if err:
    print("STDERR:", err)

print("\n=== ALL API ROUTES ===")
out2, err2 = run(f"cd {base_path} && php artisan route:list --path=api 2>&1")
print(out2)
if err2:
    print("STDERR:", err2)

client.close()
print("\nDone.")
